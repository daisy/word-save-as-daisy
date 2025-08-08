﻿using System;
using System.Collections.Generic;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2; 

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    /// <summary>
    /// Abstract class for pipeline 2 scripts launch
    /// </summary>
    public abstract class Pipeline2Script : Script
    {
        private ConverterSettings _settings = ConverterSettings.Instance;

        protected Pipeline2Script(IConversionEventsHandler e)
            : base(e) { }

        //IntPtr currentJob = IntPtr.Zero;

        /// <summary>
        /// executes script using jnet bridge
        /// </summary>
        /// <param name="inputPath">path to use for script main input</param>
        /// <param name="isQuite"></param>
        public override void ExecuteScript(string input)
        {
            ScriptRunner runner;
            bool useDAISYPipelineApp = _settings.UseDAISYPipelineApp;
            try {
                runner = useDAISYPipelineApp ? AppRunner.GetInstance(EventsHandler) : JNIRunner.GetInstance(EventsHandler);
            }
            catch (System.Exception ex) {
                EventsHandler.onPostProcessingError(
                    new Exception("An error occurred while launching the pipeline, fall back to / retry embedded engine", ex)
                );
                runner = JNIRunner.GetInstance(EventsHandler);
                useDAISYPipelineApp = false;
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
                // NP 2025/08/08 : Pipeline app requires file uris, while JNIRunner/SimpleAPI requires system file path
                if(useDAISYPipelineApp && v.Value.ParameterData is PathData path) {
                    parameters[v.Value.Name] = new Uri(path.Value.ToString()).AbsoluteUri;
                } else {
                    parameters[v.Value.Name] = v.Value.Value;
                }
            }

            runner.StartJob(Name, parameters, Parameters["output"].Value.ToString());
        }
    }
}
