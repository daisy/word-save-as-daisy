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

            JNIRunner.GetInstance(EventsHandler).StartJob(Name, parameters);

        }
    }
}
