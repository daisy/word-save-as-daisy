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
            if (!ConverterHelper.PipelineAppIsInstalled()) {
                throw new InvalidOperationException("DAISY Pipeline application is not installed.");
            }
            // Check if the app is running
            // Recherche tous les processus dont le nom correspond à "DAISY Pipeline"
            // (sans l'extension .exe, car ProcessName ne la contient jamais)
            var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ConverterHelper.PipelineAppPath));
            if(processes.Length == 0) {
                // If not running, start the app
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.PipelineAppPath,
                    Arguments = "--bg --hidden",
                    UseShellExecute = true, // Important pour détacher du parent
                    CreateNoWindow = true,  // Pas de fenêtre de console
                    WindowStyle = ProcessWindowStyle.Hidden // Optionnel, pour cacher la fenêtre si c'est une GUI
                };
                Process.Start(startInfo);
            }
            int attempt = 10;
            do {
                attempt--;
                try {
                    _webservice = getAppWebservice();
                    AliveData data = _webservice.Alive().Result;
                    if(data == null || !data.Alive) {
                        _webservice = null;
                    }
                } catch (AggregateException e) {
                    _webservice = null;
                    // If the app is not ready yet, wait a bit
                    System.Threading.Thread.Sleep(1000);
                }
                catch (Exception e) {
                    _webservice = null;
                    // If the app is not ready yet, wait a bit
                    System.Threading.Thread.Sleep(1000);
                }
            } while(_webservice == null && attempt > 0);
            if(_webservice == null) {
                throw new InvalidOperationException("Could not connect to DAISY Pipeline app webservice after 10 attempts");
            }
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

        public static new ScriptRunner GetInstance(IConversionEventsHandler events = null) {
            return new AppRunner(events);
        }
        public override void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "") {
            
            try {
                JobData data = _webservice.LaunchJob(scriptName, options).Result;
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
    }
}
