using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Allow to use step declared in the pipeline 2 as scripts
    /// </summary>
    public abstract class Pipeline2Step : Script
    {

        public override void ExecuteScript(string inputPath, bool isQuite)
        {
            Pipeline2 pipeline = Pipeline2.Instance;

            if (this.OnPipelineError != null)
            {
                Pipeline2.Instance.SetPipelineErrorListener((message) => {
                    this.OnPipelineError(message);
                });
            }
            if (this.OnPipelineOutput != null)
            {
                Pipeline2.Instance.SetPipelineOutputListener((message) => {
                    this.OnPipelineOutput(message);
                });
            }

            Dictionary<string, string> inputs = new Dictionary<string, string>();
            Dictionary<string, object> options = new Dictionary<string, object>();
            Dictionary<string, string> outputs = new Dictionary<string, string>();

            foreach (KeyValuePair<string,ScriptParameter> item in Parameters)
            {
                object value = item.Value.ParameterValue;
                if(value != null && value.ToString() != "")
                {
                    // Convert path to pipeline compatible URI
                    if (item.Value.ParameterDataType is PathDataType)
                    {
                        PathDataType temp = (PathDataType)(item.Value.ParameterDataType);
                        value = new Uri(value.ToString()).AbsoluteUri;//.Replace("file:///", "file:/");
                        if( temp.IsFileOrDirectory.Equals(PathDataType.FileOrDirectory.Directory)
                            && !value.ToString().EndsWith("/")
                        ) {
                            value += "/";
                        }
                    }
                    switch (item.Value.Direction)
                    {
                        case ScriptParameter.ParameterDirection.Input:
                            inputs.Add(item.Value.Name, value.ToString());
                            break;
                        case ScriptParameter.ParameterDirection.Output:
                            outputs.Add(item.Value.Name, value.ToString());
                            break;
                        case ScriptParameter.ParameterDirection.Option:
                            options.Add(item.Value.Name, value);
                            break;
                    }
                }
                
            }

            Pipeline2.Instance.runXprocStep(NameOrURI,inputs,outputs,options);
        }
    }
}
