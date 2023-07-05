using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Daisy.SaveAsDAISY.Conversion.Events;
using org.daisy.jnet;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Abstract class for pipeline 2 scripts launch
    /// </summary>
    public abstract class Pipeline2Script : Script {

        protected Pipeline2Script(IConversionEventsHandler e) : base(e) { }

        //IntPtr currentJob = IntPtr.Zero;

        /// <summary>
        /// executes script using jnet bridge
        /// </summary>
        /// <param name="inputPath">path to use for script main input</param>
        /// <param name="isQuite"></param>
        public override void ExecuteScript(string inputPath, bool isQuite) {

            Pipeline2 pipeline = Pipeline2.Instance;
            Pipeline2.Instance.SetPipelineErrorListener((message) => {
                this.EventsHandler.OnError(message);
            });
            Pipeline2.Instance.SetPipelineOutputListener((message) => {
                this.EventsHandler.onFeedbackMessageReceived(this, new DaisyEventArgs(message));
                this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs(message));
            });
            if (Parameters.ContainsKey("input") && (string)Parameters["input"].ParameterValue == "")
            {
                Parameters["input"].ParameterValue = inputPath;
            }
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            foreach(KeyValuePair<string, ScriptParameter> v in Parameters)
            {
                // avoid passing empty values
                if(
                    !v.Value.IsParameterRequired &&
                    (
                        v.Value.ParameterDataType is StringDataType
                        || v.Value.ParameterDataType is PathDataType
                    ) &&
                    "" == (string)v.Value.ParameterValue
                ) {
                    continue;
                }
                parameters[v.Value.Name] = v.Value.ParameterValue;
            }
            IntPtr currentJob = Pipeline2.Instance.Start(
                Name,
                parameters
            );

            if (currentJob != IntPtr.Zero) {
                //IntPtr context = pipeline.getContext(currentJob);
                IntPtr monitor = pipeline.getMonitor(currentJob);
                IntPtr messages = pipeline.getMessageAccessor(monitor);
                bool checkStatus = true;
                List<string> errors;
                while (checkStatus) {
                    foreach (string message in pipeline.getNewMessages()) {
                        this.EventsHandler.onFeedbackMessageReceived(this, new DaisyEventArgs(message));
                    }
                    //Console.WriteLine(pipeline.getProgress(messages));
                    // TODO need to get jobs log
                    //Console.WriteLine("checking status");
                    switch (pipeline.getStatus(currentJob)) {
                        case Pipeline2.JobStatus.IDLE:
                            break;
                        case Pipeline2.JobStatus.RUNNING:
                            break;
                        case Pipeline2.JobStatus.SUCCESS:
                            checkStatus = false;
                            break;
                        case Pipeline2.JobStatus.ERROR:
                            errors = pipeline.getErros(currentJob);
                            throw new Exception(this.NiceName + " conversion job has finished in error :\r\n" + string.Join("\r\n", errors));
                        case Pipeline2.JobStatus.FAIL:
                            // open jobs folder
                            errors = pipeline.getErros(currentJob);
                            throw new Exception(
                               this.NiceName + " conversion job failed :\r\n" + string.Join("\r\n",errors)
                            );
                        default:
                            break;
                    }
                    System.Threading.Thread.Sleep(1000);
#if DEBUG
                // Kill the instance and running jvm for pipeline debugging
                try {
                    //Pipeline2.KillInstance();
                } catch (Exception e) {
                    this.EventsHandler.OnError(e.Message);
                }
                //
#else
                    if (!isQuite && !string.IsNullOrEmpty(output))
                        System.Diagnostics.Process.Start(output);
#endif
                }
            } else {
                throw new Exception("An unknown error occured while launching the script " + this.Name + " with the parameters " +
                    this.Parameters.Aggregate(
                        "",
                        (result, keyvalue) => result + keyvalue.Value.Name + "=" + keyvalue.Value.ParameterValue.ToString() +"\r\n"
                    )
                );
            }

        }
    }
}
