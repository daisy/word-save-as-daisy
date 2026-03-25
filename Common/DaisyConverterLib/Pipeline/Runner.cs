using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    /// <summary>
    /// Run pipeline actions
    /// </summary>
    public abstract class Runner
    {
        /// <summary>
        /// Return an instance of the script runner. <br/>
        /// To be overridden by the concrete implementation to return an instance of the script runner.
        /// </summary>
        /// <param name="events"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static Runner GetInstance(IConversionEventsHandler events = null) { throw new NotImplementedException();  }

        public abstract List<ScriptDefinition> GetAvailableScripts(bool refresh = false);


        public abstract List<EngineProperty> GetSettableProperties();
        
        public abstract List<Datatype> GetDatatypes();

        /// <summary>
        /// Start a job using the script name and options
        /// </summary>
        /// <param name="scriptName"></param>
        /// <param name="options"></param>
        public abstract void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "");

        
        
    }
}
