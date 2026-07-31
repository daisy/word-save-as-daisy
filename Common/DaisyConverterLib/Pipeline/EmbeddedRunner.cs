using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    /// <summary>
    /// Alternative runner implementation using a swing UI defined in the embedded pipeline engine.
    /// This runner is launched similarly as the JNIWrapper but through the "pipeline2.bat" script with arguments "ui" and the script name or the command to run.
    /// </summary>
    public class EmbeddedRunner : Runner
    {

        private IConversionEventsHandler events;
        private EmbeddedRunner(IConversionEventsHandler events = null)
        {
            this.events = events ?? new SilentEventsHandler();
        }

        private static EmbeddedRunner instance = null;

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
                        instance = new EmbeddedRunner(events);
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

        #region Folders used by the embedded engine
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
            var exporter = LaunchEmbeddedEngine("descriptors", new Dictionary<string, object>() { { "result", descriptorsDirectory } });
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
            Process conversion = LaunchEmbeddedEngine(scriptName, options);
            conversion.OutputDataReceived += (sender, e) =>
            {
                events?.onFeedbackMessageReceived(sender, new DaisyEventArgs(e.Data));
            };
            conversion.ErrorDataReceived += (sender,e) => events?.onFeedbackMessageReceived(sender, new DaisyEventArgs(e.Data));
            conversion.Start();
            conversion.BeginOutputReadLine();
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

        private static string getAsFileURI(string path)
        {
            return Uri.EscapeUriString(
                new Uri("file:/" + path.Replace("\\", "/")).ToString().Replace(":///", ":/")
            );
        }

        public static Process LaunchEmbeddedEngine(string scriptNameOrCommand, Dictionary<string, object> options = null)
        {
            if (!Directory.Exists(Path.Combine(ConverterHelper.EmbeddedEnginePath)))
            {
                throw new FileNotFoundException($"The embedded pipeline is not found at : {ConverterHelper.EmbeddedEnginePath}");
            }

            List<string> JavaOptions = new List<string>
            {
                "--add-opens=java.base/java.security=ALL-UNNAMED",
                "--add-opens=java.base/java.net=ALL-UNNAMED",
                "--add-opens=java.base/java.lang=ALL-UNNAMED",
                "--add-opens=java.base/java.util=ALL-UNNAMED",
                "--add-opens=java.naming/javax.naming.spi=ALL-UNNAMED",
                "--add-opens=java.rmi/sun.rmi.transport.tcp=ALL-UNNAMED",
                "--add-exports=java.base/sun.net.www.protocol.http=ALL-UNNAMED",
                "--add-exports=java.base/sun.net.www.protocol.https=ALL-UNNAMED",
                "--add-exports=java.base/sun.net.www.protocol.jar=ALL-UNNAMED",
                "--add-exports=jdk.xml.dom/org.w3c.dom.html=ALL-UNNAMED",
                "--add-exports=jdk.naming.rmi/com.sun.jndi.url.rmi=ALL-UNNAMED",
#if DEBUG
                "-Xdebug",
                "-XX:+CreateMinidumpOnCrash"
#endif
            };


            // use jnet to execute the conversion on the inputPath
            // check for arch specific jre based on how the jre folders are created by the daisy/pipeline-assembly project (lite-bridge version)
            string arch = (IntPtr.Size * 8).ToString();
            string jrePath = Directory.Exists(ConverterHelper.EmbeddedEnginePath + @"\jre" + arch)
                ? ConverterHelper.EmbeddedEnginePath + @"\jre" + (IntPtr.Size * 8).ToString()
                : ConverterHelper.EmbeddedEnginePath + @"\jre";
            string javaexe = Path.Combine(jrePath, "bin", "java.exe");

#if DEBUG
            if (File.Exists(Path.Combine(jrePath, "release")) &&
                    File.ReadAllText(Path.Combine(jrePath, "release")).Contains("jdk.jdwp.agent"))
            {
                JavaOptions.Add("-agentlib:jdwp=transport=dt_socket,server=y,suspend=n,address=5005");
            }
#endif

            Dictionary<string, string> SystemProps = new Dictionary<string, string>
            {
                {
                    "-Dorg.daisy.pipeline.properties",
                    "\"" + Path.Combine(ConverterHelper.EmbeddedEnginePath, "etc", "pipeline.properties") + "\""
                },
                // Logback configuration file
                {
                    "-Dlogback.configurationFile",
                    getAsFileURI(Path.Combine(ConverterHelper.EmbeddedEnginePath, "etc", "logback.xml"))
                },
                // Workaround for encoding bugs on Windows
                { "-Dfile.encoding", "UTF8" },
                // to make ${org.daisy.pipeline.data}, ${org.daisy.pipeline.logdir} and ${org.daisy.pipeline.mode}
                // available in config-logback.xml and felix.properties
                // note that config-logback.xml is the only place where ${org.daisy.pipeline.mode} is used
                { "-Dorg.daisy.pipeline.data", $"\"{AppDataFolder}\"" },
                { "-Dorg.daisy.pipeline.logdir",  $"\"{LogsFolder.Replace("\\", "/")}\"" },
                //{ "-Dorg.daisy.pipeline.mode", "cli" }
            };

            if (ConverterSettings.Instance.MistralApiKey != string.Empty)
            {
                //optionsString += " -Dorg.daisy.pipeline.ocr.mistral.apikey \"" + ConverterSettings.Instance.MistralApiKey + "\"";
                if (SystemProps.ContainsKey("-Dorg.daisy.pipeline.ocr.mistral.apikey"))
                {
                    SystemProps["-Dorg.daisy.pipeline.ocr.mistral.apikey"] = ConverterSettings.Instance.MistralApiKey;
                }
                else
                {
                    SystemProps.Add("-Dorg.daisy.pipeline.ocr.mistral.apikey", ConverterSettings.Instance.MistralApiKey);
                }
            }

            string ClassPath = Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "common") + "\\*"
                + Path.PathSeparator.ToString() + Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api")
                + Path.PathSeparator.ToString() + Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api", "ui")
                + Path.PathSeparator.ToString() + Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api", "xml")
                + Path.PathSeparator.ToString() + Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api", "api");

            List<string> _options = new List<string>();
            _options = JavaOptions
                .Concat(
                    SystemProps.Aggregate(
                        new List<string>(),
                        (opts, opt) =>
                        {
                            opts.Add(opt.Key.ToString() + "=" + opt.Value.ToString());
                            return opts;
                        }
                    )
                )
                .ToList();
            _options.Add(" -classpath \"" + ClassPath + "\"");

            string mainClass = "GraphicalInterface"; // mode UI - swing-based progress dialog, made from AI port of JNI wrapper
            //string mainClass = "CommandLineInterface"; // mode CLI

            string scriptOrCommandArguments = string.Join(
                " ",
                options
                    .Where(kv => kv.Value != null)
                    .Select(kv => $"--{kv.Key} \"{(kv.Value is bool ? kv.Value.ToString().ToLower() : kv.Value)}\"")
            );

            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = javaexe,
                Arguments = $"{string.Join(" ", _options)} {mainClass} {scriptNameOrCommand} {scriptOrCommandArguments}",
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            };
            return new Process()
            {
                StartInfo = startInfo
            };
        }

        
    }
}
