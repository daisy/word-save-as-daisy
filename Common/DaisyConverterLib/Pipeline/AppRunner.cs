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
    /// Running script through the DAISY Pipeline app engine
    /// </summary>
    public class AppRunner : ScriptRunner
    {
        private Webservice _webservice;
        private IConversionEventsHandler events;
        private AppRunner(IConversionEventsHandler events = null)
        {
            this.events = events ?? new SilentEventsHandler();
            LaunchApp();
            //EnsureWebserviceAlive();
        }

        

        private static readonly Regex SeekWebserviceHost = new Regex(@"[""\']?host[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);
        private static readonly Regex SeekWebservicePort = new Regex(@"[""\']?port[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);
        private static readonly Regex SeekWebservicePath = new Regex(@"[""\']?path[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);
        private Webservice getAppWebservice()
        {
            string settingsData = File.ReadAllText(Path.Combine(ConverterHelper.PipelineAppDataPath, "settings.json"));
            string host = SeekWebserviceHost.Match(settingsData).Groups[1].Value;
            string port = SeekWebservicePort.Match(settingsData).Groups[1].Value;
            string path = SeekWebservicePath.Match(settingsData).Groups[1].Value;
            return new Webservice(host, int.Parse(port), path);
        }

        private static AppRunner instance = null;

        // for thread safety
        private static readonly object padlock = new object();

        public static new ScriptRunner GetInstance(IConversionEventsHandler events = null) {
            lock (padlock) {
                if (instance == null) {
                    try {
                        instance = new AppRunner(events);
                    }
                    catch (Exception ex) {
                        throw new Exception("An error occured while launching or connecting to DAISY Pipeline App", ex);
                    }
                } else {
                    instance.LaunchApp();
                }
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
                    events.onProgressMessageReceived(this, new DaisyEventArgs(
                        DateTimeOffset.FromUnixTimeMilliseconds(message.Timestamp).DateTime.ToString("yyyy-MM-dd-HH:mm:ss.fff") + " - " + message.Content));
                }
            }
        }

        public override void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "") {
            
            try {
                LaunchApp(); // ensure the app is running if it as been closed
                EnsureWebserviceAlive(); // ensure webservice is alive
                JobData data = _webservice.LaunchJob(scriptName, options);
                List<JobStatus> running = new List<JobStatus>() {
                    JobStatus.Idle, JobStatus.Running
                };
                Uri outputFolder = new Uri(outputPath != "" ? outputPath
                    : Path.GetTempFileName()
                );
                do {
                    Thread.Sleep(1000);
                    data = _webservice.CheckJobUpdate(data).Result;
                    switch (data.Status) {
                        case JobStatus.Idle:
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
                } while (data.Status != null && running.Contains(data.Status.Value));
                
            }
            catch(JobException) {
                throw;
            }
            catch (AggregateException e) {
                events.OnConversionError(new Exception("An error occured during the job launch or its monitoring", e));
            }
        }

        /// <summary>
        /// Check if the app is installed and launch it if it is not already running.
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        private void LaunchApp()
        {
            if (!ConverterHelper.PipelineAppIsInstalled()) {
                throw new InvalidOperationException("DAISY Pipeline application is not installed.");
            }
            // Check if the app is running
            var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ConverterHelper.PipelineAppPath));
            if (processes.Length == 0) {
                events.onProgressMessageReceived(this, new DaisyEventArgs("Starting DAISY Pipeline app..."));
                // If not running, start the app
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.PipelineAppPath,
                    Arguments = "--bg --hidden",
                    UseShellExecute = true, 
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                var appLaunched = Process.Start(startInfo);
                // Wait a second that the electorn app context is initialized
                if (appLaunched.WaitForExit(250) && appLaunched.ExitCode != 0) {
                    throw new InvalidOperationException("Could not launch DAISY Pipeline app, the process has exited with error code " + appLaunched.ExitCode);
                }
            }
        }

        /// <summary>
        /// Ensure webservice is (still) alive
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        private void EnsureWebserviceAlive()
        {
            int attempt = 20;
            do {
                events.onProgressMessageReceived(this, new DaisyEventArgs($"Checking engine web service status ({attempt} attempts remaining)..."));
                attempt--;
                try {
                    _webservice = getAppWebservice();
                    AliveData data = _webservice.Alive();
                    if (data == null || !data.Alive) {
                        _webservice = null;
                    }
                }
                catch (AggregateException e) {
                    _webservice = null;
                    System.Threading.Thread.Sleep(1000);
                }
                catch (Exception e) {
                    _webservice = null;
                    System.Threading.Thread.Sleep(1000);
                }
            } while (_webservice == null && attempt > 0);
            if (_webservice == null) {
                //events.OnConversionError(new InvalidOperationException("Could not connect to DAISY Pipeline app webservice after 10 attempts"));
                throw new InvalidOperationException("Could not connect to DAISY Pipeline app webservice after 10 attempts");
            }
        }

        /// <summary>
        /// Open the DAISY Pipeline app "Browse voices" settings.
        /// </summary>
        public void BrowseVoices() {
            LaunchApp();
            var startInfo = new ProcessStartInfo
            {
                FileName = ConverterHelper.PipelineAppPath,
                Arguments = "browse-voices",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(startInfo);
        }

        /// <summary>
        /// Open the DAISY Pipeline app "Preferred voices" settings.
        /// </summary>
        public void PreferredVoices() {
            LaunchApp();
            var startInfo = new ProcessStartInfo
            {
                FileName = ConverterHelper.PipelineAppPath,
                Arguments = "preferred-voices",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(startInfo);
        }

        /// <summary>
        /// Open the DAISY Pipeline app "Engines" settings.
        /// </summary>
        public void TTSEngines() {
            LaunchApp();
            var startInfo = new ProcessStartInfo
            {
                FileName = ConverterHelper.PipelineAppPath,
                Arguments = "engines",
                UseShellExecute = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(startInfo);
        }
    }
}
