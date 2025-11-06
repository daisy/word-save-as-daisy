using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public class Engine
    {
        private static readonly Regex SeekWebserviceHostJSON = new Regex(@"[""\']?host[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);
        private static readonly Regex SeekWebservicePortJSON = new Regex(@"[""\']?port[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);
        private static readonly Regex SeekWebservicePathJSON = new Regex(@"[""\']?path[""']?\s*:\s*[""']?([^""',]+)[""']?\s*\}?,?", RegexOptions.Compiled);


        private static Webservice _appWebservice = null;
        public static Webservice getAppWebservice()
        {
            if (_appWebservice != null) {
                return _appWebservice;
            }
            string settingsData = File.ReadAllText(System.IO.Path.Combine(ConverterHelper.PipelineAppDataPath, "settings.json"));
            string host = SeekWebserviceHostJSON.Match(settingsData).Groups[1].Value;
            string port = SeekWebservicePortJSON.Match(settingsData).Groups[1].Value;
            string path = SeekWebservicePathJSON.Match(settingsData).Groups[1].Value;
            _appWebservice = new Webservice(host, int.Parse(port), path);
            return _appWebservice;
        }


        private static readonly Regex SeekWebserviceHostXML = new Regex(
            @"org\.daisy\.pipeline\.ws\.host=\s*[""']?([^""',\r\n]+)[""']?\r?\n",
            RegexOptions.Compiled
        );
        private static readonly Regex SeekWebservicePortXML = new Regex(
            @"org\.daisy\.pipeline\.ws\.port=\s*[""']?([^""',\r\n]+)[""']?\r?\n",
            RegexOptions.Compiled
        );
        private static readonly Regex SeekWebservicePathXML = new Regex(
            @"org\.daisy\.pipeline\.ws\.path=\s*[""']?([^""',\r\n]+)[""']?\r?\n",
            RegexOptions.Compiled
        );

        private static Webservice _embeddedWebservice = null;

        public static Webservice getEmbeddedWebservice()
        {
            if(_embeddedWebservice != null) {
                return _embeddedWebservice;
            }
            string pipelineSettingsPath = System.IO.Path.Combine(
                ConverterHelper.EmbeddedEnginePath,
                "etc", "pipeline.properties"
            );
            string settingsData = File.ReadAllText(pipelineSettingsPath);
            string host = SeekWebserviceHostXML.Match(settingsData).Groups[1].Value;
            string port = SeekWebservicePortXML.Match(settingsData).Groups[1].Value;
            string path = SeekWebservicePathXML.Match(settingsData).Groups[1].Value;
            _embeddedWebservice = new Webservice(host, int.Parse(port), path);
            return _embeddedWebservice;
        }

        public static void StartDAISYPipelineApp(IConversionEventsHandler events = null)
        {
            if (!ConverterHelper.PipelineAppIsInstalled()) {
                throw new InvalidOperationException("DAISY Pipeline application is not installed.");
            }
            // Check if the app is running
            var processes = Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(ConverterHelper.PipelineAppPath));
            if (processes.Length == 0) {
                _appWebservice = null; // reset webservice instance
                events?.onProgressMessageReceived(null, new DaisyEventArgs("Starting DAISY Pipeline app..."));
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
                //return appLaunched;
            } // else return processes[0];

        }


        public static void StartEmbeddedEngine(IConversionEventsHandler events = null)
        {
            if (!ConverterHelper.PipelineIsInstalled()) {
                throw new InvalidOperationException("Embedded engine was not found");
            }
            // Check if a java app from SaveAsDAISY is running
            var allProcesses = Process.GetProcessesByName("java")
                .Where((p) =>
                {
                    return p.MainModule.FileName.StartsWith(ConverterHelper.EmbeddedEnginePath);
                }).ToArray();
            if (allProcesses.Length == 0) {
                _embeddedWebservice = null; // reset webservice instance
                // If not running, start the embedded engine using the batch script
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.EmbeddedEngineLauncherPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                var appLaunched = Process.Start(startInfo);
                if (appLaunched.WaitForExit(250) && appLaunched.ExitCode != 0) {
                    throw new InvalidOperationException("Could not launch DAISY Pipeline app, the process has exited with error code " + appLaunched.ExitCode);
                }
                //return appLaunched;
            }// else return allProcesses[0];
        }

        public static void StopEmbeddedEngine(IConversionEventsHandler events = null)
        {
            _embeddedWebservice = null; // reset webservice instance
            var allProcesses = Process.GetProcessesByName("java")
                .Where((p) =>
                {
                    return p.MainModule.FileName.StartsWith(ConverterHelper.EmbeddedEnginePath);
                }).ToArray();
            if (allProcesses.Length > 0) {
                allProcesses[0].Kill();
            }
        }

        /// <summary>
        /// Open the DAISY Pipeline app "Browse voices" settings.
        /// </summary>
        public static void BrowseVoices()
        {

            Engine.StartDAISYPipelineApp();
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
        public static void PreferredVoices()
        {
            Engine.StartDAISYPipelineApp();
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
        public static void TTSEngines()
        {
            Engine.StartDAISYPipelineApp();
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
