using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PipelineConnector {
    class Program {
        static void Main(string[] args) {
            Pipeline2 pipeline = Pipeline2.Instance;
            string testOutput = "";
            pipeline.SetPipelineOutputListener(
                (message) => {
                    testOutput += message + "\r\n";
                });
            pipeline.SetPipelineErrorListener(
                (message) => {
                    testOutput += message + "\r\n";
                });

            //IntPtr job1 = pipeline.Start("dtbook-validator");
            /*IntPtr job2 = pipeline.Start("dtbook-validator", new Dictionary<string, string>() {
                { "input-dtbook", @"C:\Users\nicol\Desktop\test2\test.xml" }
            });*/
            IntPtr job2 = pipeline.Start("daisy3-to-epub3", new Dictionary<string, string>() {
                { "source", @"C:\Users\nicol\Desktop\test2\speechgen.opf" },
                {"output-dir", @"C:\Users\nicol\Desktop\test\daisy3-to-epub3\" }
            });
            //
            if (job2 != IntPtr.Zero) {
                IntPtr context = pipeline.getContext(job2);
                IntPtr monitor = pipeline.getMonitor(context);
                IntPtr messages = pipeline.getMessageAccessor(monitor);
                bool checkStatus = true;
                while (checkStatus) {
                    Console.WriteLine(pipeline.getProgress(messages));
                    foreach (string message in pipeline.getMessages(messages)) {
                        Console.WriteLine(message);
                    }
                    //Console.WriteLine("checking status");
                    switch (pipeline.getStatus(job2)) {
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
                    System.Threading.Thread.Sleep(100);
                }
                
            }

            pipeline.SetPipelineOutputListener(null);
            pipeline.SetPipelineErrorListener(null);

            //Console.WriteLine(pipeline.getLastOutput());
            Console.WriteLine(testOutput);
            // */
            return;
        }
    }
}
