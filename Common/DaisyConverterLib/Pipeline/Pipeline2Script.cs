using System;
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
        public override void ExecuteScript(string inputPath, bool isQuite)
        {
            
            if (Parameters.ContainsKey("input") && (string)Parameters["input"].ParameterValue == "")
            {
                Parameters["input"].ParameterValue = inputPath;
            }
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            foreach (KeyValuePair<string, ScriptParameter> v in Parameters)
            {
                // avoid passing empty values
                if (
                    !v.Value.IsParameterRequired
                    && (
                        v.Value.ParameterDataType is StringDataType
                        || v.Value.ParameterDataType is PathDataType
                    )
                    && "" == (string)v.Value.ParameterValue
                )
                {
                    continue;
                }
                parameters[v.Value.Name] = v.Value.ParameterValue;
            }

            ScriptRunner runner;
            try {
                runner = _settings.UseDAISYPipelineApp ? AppRunner.GetInstance(EventsHandler) : JNIRunner.GetInstance(EventsHandler);
            } catch (System.Exception ex) {
                EventsHandler.onPostProcessingError(
                    new Exception("An error occurred while launching the pipeline, fall back to / retry embedded engine", ex)
                );
                runner = JNIRunner.GetInstance(EventsHandler);
            }
            runner.StartJob(Name, parameters, Parameters["output"].ParameterValue.ToString());

        }
    }
}
