using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// pipeline script interface
    /// 
    /// </summary>
    public abstract class Script {

        protected IConversionEventsHandler eventsHandler;

        public Script(IConversionEventsHandler eventsHandler = null)
        {
            this.eventsHandler = eventsHandler ?? new SilentEventsHandler();
        }

        public IConversionEventsHandler EventsHandler { get { return eventsHandler; } set { eventsHandler = value; } }

        public delegate void PipelineOutputListener(object sender, EventArgs e);
        protected event PipelineOutputListener onPipelineOutput;
        public void setPipelineOutputListener(PipelineOutputListener onPipelineOutput) {
            this.onPipelineOutput = onPipelineOutput;
        }

        protected PipelineOutputListener OnPipelineOutput {
            get => this.onPipelineOutput;
        }


        public delegate void PipelineErrorListener(object sender, EventArgs e);
        protected event PipelineErrorListener onPipelineError;
        public void setPipelineErrorListener(PipelineErrorListener onPipelineError) {
            this.onPipelineError = onPipelineError;
        }
        protected PipelineErrorListener OnPipelineError {
            get => this.onPipelineError;
        }

        public delegate void PipelineProgressListener(object sender, EventArgs e);
        protected event PipelineProgressListener onPipelineProgress;
        public void setPipelineProgressListener(PipelineProgressListener onPipelineProgress) {
            this.onPipelineProgress = onPipelineProgress;
        }
        protected PipelineProgressListener OnPipelineProgress {
            get => this.onPipelineProgress;
        }




        protected Dictionary<string, ScriptParameter> _parameters;
        /// <summary>
        /// List of parameters available in script
        /// </summary>
        public Dictionary<string, ScriptParameter> Parameters {
            get { return _parameters; }
        }

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
        public string Description
        {
            get { return name; }
        }

        /// <summary>
        /// Script output path, to be extracted from parameters in runners
        /// </summary>
        public string output = string.Empty;

        /// <summary>
        /// Executes script normal.
        /// </summary>
        /// <param name="inputPath"></param>
        public void ExecuteScript(string inputPath) {
            ExecuteScript(inputPath, false);
        }

        /// <summary>
        ///  executes script
        /// </summary>
        public abstract void ExecuteScript(string inputPath, bool isQuite);

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

    }
}
