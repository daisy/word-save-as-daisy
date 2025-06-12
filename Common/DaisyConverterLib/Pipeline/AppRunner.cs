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
            EnsureWebserviceAlive();
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
                }
                return instance;
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
                            List<Message> messages = data.Messages;
                            break;
                        case JobStatus.Success:
                            data = _webservice.DownloadResults(data, outputFolder.LocalPath).Result;
                            events.onProgressMessageReceived(this, new DaisyEventArgs(scriptName + " finished successfully"));
                            break;
                        case JobStatus.Error:
                            data = _webservice.DownloadResults(data, outputFolder.LocalPath).Result;
                            events.OnConversionError(new Exception("The job finished in error, please consult the log file at " + data.Log));
                            break;
                        default:

                            break;
                    }
                } while (data.Status != null && running.Contains(data.Status.Value));
                
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
                Process.Start(startInfo);
                // Wait a second that the electorn app context is initialized
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// Ensure webservice is (still) alive
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        private void EnsureWebserviceAlive()
        {
            int attempt = 10;
            do {
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
