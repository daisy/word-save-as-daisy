using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public class JNIWrapperRunner : ScriptRunner
    {

        private IConversionEventsHandler events;
        private JNIWrapperRunner(IConversionEventsHandler events = null)
        {
            this.events = events ?? new SilentEventsHandler();
        }

        private static JNIWrapperRunner instance = null;

        // for thread safety
        private static readonly object padlock = new object();


        public static new ScriptRunner GetInstance(IConversionEventsHandler events = null)
        {
            lock (padlock)
            {
                if (instance == null)
                {
                    try
                    {
                        instance = new JNIWrapperRunner(events);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(
                            "An error occured while launching or connecting to DAISY Pipeline App",
                            ex
                        );
                    }
                }

                return instance;
            }
        }

        public override void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "")
        {
            // NP : I'm replacing the jni runner code by a wrapper application
            // Embedding the wrapper directly in the addin tends to cause a lot of issue with service loading
            // (trying to reuse an instance of the SimpleAPI leads to a "classLoader is null" exception)
            // i'm thinking that using a small JNI wrapper app with WPF might be a good solution

            events?.onFeedbackMessageReceived(this, new DaisyEventArgs($"Starting conversion with script {scriptName} using the jni wrapper"));
            string appPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ConversionApp.exe");
            string optionsString = scriptName + " " + string.Join(
                " ",
                options
                    .Where(kv => kv.Value != null)
                    .Select(kv => $"--{kv.Key} \"{(kv.Value is bool ? kv.Value.ToString().ToLower() : kv.Value)}\"")
            );
            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = appPath,
                Arguments = optionsString,
                UseShellExecute = false,
                RedirectStandardError = true,
//                CreateNoWindow = true,
            };

            Process conversion = new Process()
            {
                StartInfo = startInfo,
            };
            conversion.ErrorDataReceived += (sender,e) => events?.onFeedbackMessageReceived(sender, new DaisyEventArgs(e.Data));
            conversion.Start();
            //conversion.BeginOutputReadLine();
            conversion.BeginErrorReadLine();
            conversion.WaitForExit();
            //events?.onProgressMessageReceived(this, new DaisyEventArgs("Conversion app exited with code " + conversion.ExitCode));
            if(conversion.ExitCode > 1)
            {
                throw new JobException($"Conversion app exited with code {conversion.ExitCode}");
            }
        }
    }
}
