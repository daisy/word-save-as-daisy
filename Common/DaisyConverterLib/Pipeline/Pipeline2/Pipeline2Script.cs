using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using org.daisy.jnet;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Abstract class for pipeline 2 scripts launch
    /// </summary>
    public abstract class Pipeline2Script : Script {

        //IntPtr currentJob = IntPtr.Zero;

        /// <summary>
        /// executes script using jnet bridge
        /// </summary>
        /// <param name="inputPath">not used for pipeline 2 script : input is a script parameter</param>
        /// <param name="isQuite"></param>
        public override void ExecuteScript(string inputPath, bool isQuite) {

            Pipeline2 pipeline = Pipeline2.Instance;
            // convert parameters list to options dictionnary
            if (this.OnPipelineError != null) {
                Pipeline2.Instance.SetPipelineErrorListener((message) => {
                    this.OnPipelineError(message);
                });
            }
            if (this.OnPipelineOutput != null) {
                Pipeline2.Instance.SetPipelineOutputListener((message) => {
                    this.OnPipelineOutput(message);
                });
            }

            IntPtr currentJob = Pipeline2.Instance.Start(
                Name,
                Parameters.ToDictionary(keyvalue => keyvalue.Value.Name, keyvalue => keyvalue.Value.ParameterValue)
            );



            if (currentJob != IntPtr.Zero) {
                IntPtr context = pipeline.getContext(currentJob);
                IntPtr monitor = pipeline.getMonitor(context);
                IntPtr messages = pipeline.getMessageAccessor(monitor);
                bool checkStatus = true;
                while (checkStatus) {
                    if(this.OnPipelineOutput != null) {
                        foreach (string message in pipeline.getLastMessages(messages)) {
                            this.OnPipelineOutput(message);
                        }
                        
                    }
                    //Console.WriteLine(pipeline.getProgress(messages));
                    
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
                            checkStatus = false;
                            break;
                        case Pipeline2.JobStatus.FAIL:
                            checkStatus = false;
                            break;
                        default:
                            break;
                    }
                    System.Threading.Thread.Sleep(1000);
#if DEBUG
                // Kill the instance and running jvm for pipeline debugging
                try {
                    //Pipeline2.KillInstance();
                } catch (Exception e) {
                    if (this.OnPipelineError != null) {
                        this.OnPipelineError(e.Message);
                    }
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
