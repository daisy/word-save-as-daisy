using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public abstract class ScriptRunner
    {
        /// <summary>
        /// Return an instance of the script runner. <br/>
        /// To be overridden by the concrete implementation to return an instance of the script runner.
        /// </summary>
        /// <param name="events"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public static ScriptRunner GetInstance(IConversionEventsHandler events = null) { throw new NotImplementedException();  }

        /// <summary>
        /// Start a job using the script name and options
        /// </summary>
        /// <param name="scriptName"></param>
        /// <param name="options"></param>
        public abstract void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "");


    }
}
