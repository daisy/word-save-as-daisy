using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public abstract class ScriptRunner
    {
        public static ScriptRunner GetInstance(IConversionEventsHandler events = null) { throw new NotImplementedException();  }

        public void StartJob(string scriptName, Dictionary<string, object> options = null) { throw new NotImplementedException(); }


    }
}
