using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime;

namespace Daisy.SaveAsDAISY.Conversion
{

    /// <summary>
    /// pipeline script interface
    /// 
    /// </summary>
    public abstract class Script
    {

        protected IConversionEventsHandler eventsHandler;

        private ConverterSettings _settings = ConverterSettings.Instance;

        public Script(IConversionEventsHandler eventsHandler = null)
        {
            this.eventsHandler = eventsHandler ?? new SilentEventsHandler();
            _parameters = new Dictionary<string, ScriptParameter>();
        }

        public IConversionEventsHandler EventsHandler { get { return eventsHandler; } set { eventsHandler = value; } }

        public delegate void PipelineOutputListener(object sender, EventArgs e);
        protected event PipelineOutputListener onPipelineOutput;
        public void setPipelineOutputListener(PipelineOutputListener onPipelineOutput)
        {
            this.onPipelineOutput = onPipelineOutput;
        }

        protected PipelineOutputListener OnPipelineOutput {
            get => this.onPipelineOutput;
        }


        public delegate void PipelineErrorListener(object sender, EventArgs e);
        protected event PipelineErrorListener onPipelineError;
        public void setPipelineErrorListener(PipelineErrorListener onPipelineError)
        {
            this.onPipelineError = onPipelineError;
        }
        protected PipelineErrorListener OnPipelineError {
            get => this.onPipelineError;
        }

        public delegate void PipelineProgressListener(object sender, EventArgs e);
        protected event PipelineProgressListener onPipelineProgress;
        public void setPipelineProgressListener(PipelineProgressListener onPipelineProgress)
        {
            this.onPipelineProgress = onPipelineProgress;
        }
        protected PipelineProgressListener OnPipelineProgress {
            get => this.onPipelineProgress;
        }


        protected Dictionary<string, ScriptParameter> _parameters;

        protected bool _alsoExportShapes = false;

        /// <summary>
        /// List of parameters available in script
        /// </summary>
        public Dictionary<string, ScriptParameter> Parameters {
            get { return _parameters; }
        }

        public List<string> ExtractedShapes { get; set; } = new List<string>();

        protected string niceName = "";

        /// <summary>
        /// Name to be displayed in UI
        /// </summary>
        public string NiceName {
            get { return niceName; }
        }

        protected string name = "";

        /// <summary>
        /// Name used in pipeline calls (_postprocess, or dtbook-to-daisy303)
        /// </summary>
        public string Name {
            get { return name; }
        }

        protected string description = "";

        /// <summary>
        /// Description of the script
        /// </summary>
        public string Description {
            get { return name; }
        }

        /// <summary>
        /// Script output path, to be extracted from parameters in runners
        /// </summary>
        public string output = string.Empty;

        /// <summary>
        /// Execute the script on the given input file path.
        /// </summary>
        /// <param name="input">input file path</param>
        public virtual void ExecuteScript(string input)
        {
            Runner runner;
            if (_settings.UseWebserviceRunner)
            {
                runner = WebserviceRunner.GetInstance(EventsHandler);
            }
            else
            {
                runner = JNIWrapperRunner.GetInstance(EventsHandler);
            }
            if (Parameters.ContainsKey("input") && (string)Parameters["input"].Value == "")
            {
                Parameters["input"].Value = input;
            }
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            foreach (KeyValuePair<string, ScriptParameter> v in Parameters)
            {

                // avoid passing empty values
                if (
                    !v.Value.IsParameterRequired
                    && (
                        v.Value.ParameterData is StringData
                        || v.Value.ParameterData is PathData
                    )
                    && "" == (string)v.Value.Value
                )
                {
                    continue;
                }
                // NP 2025/08/08 : Webservice requires file uris, while JNIRunner/SimpleAPI requires system file path
                if (v.Value.ParameterData is PathData path && runner is WebserviceRunner)
                {
                    parameters[v.Value.Name] = new Uri(path.Value.ToString()).AbsoluteUri;
                }
                else
                {
                    parameters[v.Value.Name] = v.Value.Value;
                }
            }

            runner.StartJob(Name, parameters, Parameters["output"].Value.ToString());

            if (_alsoExportShapes && ExtractedShapes.Count > 0)
            {
                foreach (string shape in ExtractedShapes)
                {
                    if (File.Exists(shape) == false)
                    {
                        // File has been already copied by the new word to dtbook script, continue parsing
                        continue;
                    }
                    try
                    {
                        File.Copy(shape, Path.Combine(Parameters["output"].Value.ToString(), Path.GetFileName(shape)), true);
                    }
                    catch (Exception ex)
                    {
                        EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Error while copying extracted shape " + shape + ": " + ex.Message));
                    }
                }
            }
        }


        /// <summary>
        /// Searches for input file in the given directory.
        /// </summary>
        /// <param name="inputDirectory"></param>
        /// <returns></returns>
        public abstract string searchInputFromDirectory(DirectoryInfo inputDirectory);

        /// <summary>
        /// Number of steps for progression
        /// </summary>
        public int StepsCount { get; protected set; } = 1;

        /// <summary>
        /// Load document parameters into the script parameters.
        /// </summary>
        /// <param name="doc"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="KeyNotFoundException"></exception>
        public void OnDocument(DocumentProperties doc)
        {
            if (doc == null || string.IsNullOrEmpty(doc.InputPath)) {
                throw new ArgumentException("DocumentProperties cannot be null or empty");
            }

            if (Parameters.ContainsKey("title"))
                Parameters["title"].Value = doc.Title;

            if (Parameters.ContainsKey("creator"))
                Parameters["creator"].Value = doc.Author;

            if (Parameters.ContainsKey("publisher"))
                Parameters["publisher"].Value = doc.Publisher;

            if (Parameters.ContainsKey("uid"))
                Parameters["uid"].Value = doc.Identifier;

            if (Parameters.ContainsKey("subject"))
                Parameters["subject"].Value = doc.Subject;

        }
    }
}
