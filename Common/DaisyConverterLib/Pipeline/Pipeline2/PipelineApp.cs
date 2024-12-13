

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;

namespace Daisy.SaveAsDAISY.Conversion
{

    public class PipelineApp
    {
        public static readonly string APPPATH = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs",
                    "pipeline-ui",
                    "DAISY Pipeline.exe"
                );
        public static bool IsInstalled()
        {
            return File.Exists(APPPATH);
        }

        public static bool IsRunning()
        {
            return Process.GetProcessesByName("DAISY Pipeline").Length > 0;
        }

        public static Webservice Start()
        {
            if(!IsInstalled()) {
                throw new FileNotFoundException("DAISY Pipeline app is not installed");
            }
            if(IsRunning()) {
                return FindWebservice();
            }
            ProcessStartInfo startInfo = new ProcessStartInfo(APPPATH, "--hidden");
            startInfo.WindowStyle = ProcessWindowStyle.Minimized;
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            Process app = Process.Start(startInfo);
            return FindWebservice();
        }

        public static Webservice FindWebservice()
        {
            int maxTry = 10;
            int countTry = 0;
            do {
                countTry = 1;
                Process[] toCheck = Process.GetProcessesByName("java");
                foreach (Process process in toCheck) {
                    try {
                        string[] commandline = Utils.GetCommandLine(process).Split(' ');
                        // Search the pipeline app embedded engine (under a pipeline-ui folder)
                        // TODO : maybe not the best way to get the correct jre
                        // Might need to check if it is near the APPPATH parent directory

                        if (commandline[0].StartsWith(Path.GetDirectoryName(APPPATH))) {
                            Webservice result = new Webservice()
                            {
                                BasePath = "/ws",
                                Host = "localhost",
                                Port = "8181"
                            };
                            string[] pipelineWebServiceSettings = commandline
                                .Where(s => s.StartsWith("-Dorg.daisy.pipeline.ws."))
                                .ToArray();
                            foreach (var prop in pipelineWebServiceSettings) {
                                string[] kv = prop.Split('=');
                                if (kv[0].EndsWith("host")) {
                                    result.Host = kv[1];
                                } else if (kv[0].EndsWith("path")) {
                                    result.BasePath = "/" + kv[1].TrimStart('/');
                                } else if (kv[0].EndsWith("port")) {
                                    result.Port = kv[1];
                                }
                            }
                            return result;
                        }

                    }
                    catch (Win32Exception ex) when ((uint)ex.ErrorCode == 0x80004005) {
                        // Intentionally empty - no security access to the process.
                    }
                    catch (InvalidOperationException) {
                        // Intentionally empty - the process exited before getting details.
                    }
                }
                Thread.Sleep(1000);
            } while(countTry < maxTry);
            return null;
        }

        protected static class Utils
        {
            public static string GetCommandLine(Process process)
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + process.Id))
                using (ManagementObjectCollection objects = searcher.Get()) {
                    return objects.Cast<ManagementBaseObject>().SingleOrDefault()?["CommandLine"]?.ToString();
                }

            }
        }

    }

    
    

}
