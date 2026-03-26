using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public class JNIWrapperRunner : Runner
    {

        private IConversionEventsHandler events;
        private JNIWrapperRunner(IConversionEventsHandler events = null)
        {
            this.events = events ?? new SilentEventsHandler();
        }

        private static JNIWrapperRunner instance = null;

        // for thread safety
        private static readonly object padlock = new object();


        public static new Runner GetInstance(IConversionEventsHandler events = null)
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

        #region Folders used by the JNI wrapper and the embedded engine
        public static string AppDataFolder
        {
            get
            {
                return Directory
                    .CreateDirectory(
                        Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                            "DAISY Pipeline 2"
                        )
                    )
                    .FullName;
            }
        }

        public static string LogsFolder
        {
            get { return Directory.CreateDirectory(Path.Combine(AppDataFolder, "log")).FullName; }
        }

        public static string JNILogsPath
        {
            get { return Directory.CreateDirectory(Path.Combine(AppDataFolder, "jni-logs")).FullName; }
        }
        #endregion

        private List<ScriptDefinition> _scripts;
        //private List<Datatype> _datatypes; // Datatype descriptors format is a bit of a mess with lots of namespaces without a real consistence in that definition
        // so i'm not doing this parsing yet (i'll port the code from the pipeline ui)
        

        public void RefreshCache()
        {
            _scripts = null;
            //_datatypes = null;
            _properties = null;
            string descriptorsDirectory = Path.Combine(ConverterSettings.ApplicationDataFolder);
            var exporter = LaunchJNIWrapper("descriptors", new Dictionary<string, object>() { { "output", descriptorsDirectory } });
            string errorData = "";
            exporter.ErrorDataReceived += (sender, e) => errorData += e.Data;
            exporter.Start();
            exporter.BeginErrorReadLine();
            exporter.WaitForExit();
            if(exporter.ExitCode != 0)
            {
                Exception e = new Exception($"JNI wrapper exited with code {exporter.ExitCode} while trying to refresh the script and properties cache", new Exception(errorData));
                AddinLogger.Error(e);
                throw e;
            }
            XmlDocument doc = new XmlDocument();
            if (File.Exists(Path.Combine(descriptorsDirectory, "scripts.xml")))
            {
                try
                {
                    doc.Load(Path.Combine(descriptorsDirectory, "scripts.xml"));
                    _scripts = ScriptDefinition.ListFromXml((XmlElement)doc.GetElementsByTagName("scripts")[0]);
                }
                catch (Exception ex)
                {
                    AddinLogger.Error(new Exception("Could not read the scripts.xml file", ex));
                    //throw new Exception("An error occured while loading descriptors from the JNI wrapper output", ex);
                    _scripts = new List<ScriptDefinition>();
                }
            }
            // _datatypes = Datatype.ListFromXml((XmlElement)doc.GetElementsByTagName("datatypes")[0]);
            //_scripts = Script.ListFromXml((XmlElement)doc.GetElementsByTagName("scripts")[0]);

        }


        public override void StartJob(string scriptName, Dictionary<string, object> options = null, string outputPath = "")
        {
            // NP : I'm replacing the jni runner code by a wrapper application
            // Embedding the wrapper directly in the addin tends to cause a lot of issue with service loading
            // (trying to reuse an instance of the SimpleAPI leads to a "classLoader is null" exception)
            // i'm thinking that using a small JNI wrapper app with WPF might be a good solution

            events?.onFeedbackMessageReceived(this, new DaisyEventArgs($"Starting conversion with script {scriptName} using the jni wrapper"));
            Process conversion = LaunchJNIWrapper(scriptName, options);
            conversion.ErrorDataReceived += (sender,e) => events?.onFeedbackMessageReceived(sender, new DaisyEventArgs(e.Data));
            conversion.Start();
            //conversion.BeginOutputReadLine();
            conversion.BeginErrorReadLine();
            conversion.WaitForExit();
            switch (conversion.ExitCode)
            {
                case 0:
                    //events?.onFeedbackMessageReceived(this, new DaisyEventArgs("Conversion completed successfully"));
                    break;
                case 1:
                    throw new OperationCanceledException("Conversion was cancelled by the user");
                default:
                    throw new JobException($"Embedded pipeline returned an error, please check the latest error logs in {JNILogsPath}");
            }
        }

        public override List<ScriptDefinition> GetAvailableScripts(bool refresh = false)
        {
            if(refresh || _scripts == null)
            {
                RefreshCache();
            }
            return _scripts;
        }

        private List<EngineProperty> _properties;
        public override List<EngineProperty> GetSettableProperties()
        {
            
            if (_properties == null)
            {
                // Read properties from the daisy-pipeline/settable-properties.xml file generated by the JNI wrapper
                string propsPath = Path.Combine(ConverterHelper.EmbeddedEnginePath, "settable-properties.xml");
                XmlDocument doc = new XmlDocument();
                if (File.Exists(propsPath))
                {
                    try
                    {
                        doc.Load(propsPath);
                        _properties = EngineProperty.ListFromXml((XmlElement)doc.GetElementsByTagName("properties")[0]).OrderBy(p => p.Name).ToList();
                    }
                    catch (Exception ex)
                    {
                        AddinLogger.Error(new Exception("Could not read the settable-properties.xml file", ex));
                        //throw new Exception("An error occured while loading descriptors from the JNI wrapper output", ex);
                        _properties = new List<EngineProperty>();
                    }
                }
            }
            return _properties;
        }

        public override List<Datatype> GetDatatypes()
        {
            throw new NotImplementedException();
        }

        public static Process LaunchJNIWrapper(string scriptNameOrCommand, Dictionary<string, object> options = null)
        {
            string appPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "JNIWrapper.exe");
            string optionsString = scriptNameOrCommand + " " + string.Join(
                " ",
                options
                    .Where(kv => kv.Value != null)
                    .Select(kv => $"--{kv.Key} \"{(kv.Value is bool ? kv.Value.ToString().ToLower() : kv.Value)}\"")
            ) + string.Join(
                " ",
                PipelineUserProperties.Instance.Items
                    .Where(p => p.Value != null && p.Value != string.Empty)
                    .Select(p => $"-D{p.Name} \"{p.Value}\"")
            );
            

            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = appPath,
                Arguments = optionsString,
                UseShellExecute = false,
                RedirectStandardError = true,
                //                CreateNoWindow = true,
            };
            return new Process()
            {
                StartInfo = startInfo
            };
        }

        
    }
}
