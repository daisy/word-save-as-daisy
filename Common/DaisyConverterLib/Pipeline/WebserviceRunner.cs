using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2
{
    /// <summary>
    /// Running script through webservice of the DAISY Pipeline app
    /// </summary>
    public class WebserviceRunner : ScriptRunner
    {
        private Webservice _webservice;
        private IConversionEventsHandler events;
        private WebserviceRunner(IConversionEventsHandler events = null)
        {
            this.events = events ?? new SilentEventsHandler();
            
        }



        private static WebserviceRunner instance = null;

        // for thread safety
        private static readonly object padlock = new object();

        public static new ScriptRunner GetInstance(IConversionEventsHandler events = null) {
            lock (padlock) {
                if (instance == null) {
                    try {
                        instance = new WebserviceRunner(events);
                    }
                    catch (Exception ex) {
                        throw new Exception("An error occured while launching or connecting to DAISY Pipeline App", ex);
                    }
                }
                // Launch the engine if not already running
                instance.LaunchEngine();
                return instance;
            }
        }
        List<Message> printed = new List<Message>();

        List<Message> messagesFlattened(List<Message> l)
        {
            List<Message> res = new List<Message>();

            if (l == null) return res;
            if (l.Count == 0) return res;
            foreach (Message m in l) {
                // 
                res.Add(new Message()
                {
                    Content = m.Content,
                    Timestamp = m.Timestamp,
                });
                if (m.Messages != null && m.Messages.Count > 0) {
                    res.AddRange(messagesFlattened(m.Messages));
                }
            }
            return res.OrderBy(m => m.Timestamp).ToList();
        }
        bool isPrinted(Message m) {
            foreach (Message pm in printed) {
                if (pm.Timestamp == m.Timestamp && pm.Content == m.Content) {
                    return true;
                }
            }
            return false;
        }

        void printMessages(List<Message> messages, IConversionEventsHandler events) {
            foreach (Message message in messagesFlattened(messages)) {
                if(isPrinted(message)) {
                    continue;
                } else {
                    printed.Add(message);
                    // DateTimeOffset.FromUnixTimeMilliseconds(message.Timestamp).DateTime.ToString("yyyy-MM-dd-HH:mm:ss.fff") + " - " + 
                    events.onProgressMessageReceived(this, new DaisyEventArgs(
                        message.Content, message.Timestamp
                    ) );
                }
            }
        }

        public override void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "") {
            
            try {
                LaunchEngine(); // ensure the app is running if it as been closed
                _webservice.WaitForActivation(events);
                JobData data = _webservice.LaunchJob(scriptName, options);
                List<JobStatus> running = new List<JobStatus>() {
                    JobStatus.Idle, JobStatus.Running
                };
                Uri outputFolder = new Uri(outputPath != "" ? outputPath
                    : Path.GetTempFileName()
                );
                long timestamp = DateTimeOffset.Now.ToUnixTimeMilliseconds();

                bool cancelRequested = false;
                do {
                    Thread.Sleep(1000);
                    data = _webservice.CheckJobUpdate(data).Result;
                    switch (data.Status) {
                        case JobStatus.Idle:
                            printMessages(new List<Message>()
                            {
                                new Message()
                                {
                                    Content = scriptName + " job is idle, waiting for activation...",
                                    Timestamp = timestamp
                                }
                            }, events);
                            break;
                        case JobStatus.Running:
                            printMessages(data.Messages, events);
                            break;
                        case JobStatus.Success:
                            data = _webservice.DownloadResults(data, outputFolder.LocalPath).Result;
                            events.onProgressMessageReceived(this, new DaisyEventArgs(scriptName + " finished successfully"));
                            break;
                        case JobStatus.Fail:
                        case JobStatus.Error:
                            printMessages(data.Messages, events);
                            data = _webservice.DownloadResults(data, outputFolder.LocalPath).Result;
                            events.OnConversionError(new Exception("The job finished in status " + data.Status + ", please consult the log file at " + data.Log));
                            throw new JobException("The job finished in " + data.Status + ", please consult the log file at " + data.Log);
                        default:

                            break;
                    }

                    cancelRequested = events?.IsCancellationRequested() ?? false;
                } while (data.Status != null && running.Contains(data.Status.Value) && !cancelRequested);
                
            }
            catch(JobException) {
                throw;
            }
            catch (JobRequestError jre)
            {
                var ex = new Exception(jre.ToString());
                events.OnConversionError(ex);
                throw ex;
            }
            catch (AggregateException e) {
                var ex = new Exception("An error occured during the job launch or its monitoring", e);
                events.OnConversionError(ex);
                throw ex;
            }
        }

        /// <summary>
        /// Check if the app is installed and launch it if it is not already running.
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void LaunchEngine()
        {
            if(ConverterSettings.Instance.UseDAISYPipelineApp == false) {
                _webservice = Engine.StartEmbeddedWebservice();
            } else {
                _webservice = Engine.StartDAISYPipelineAppWebservice();
            }
           
        }


        
    }
}
