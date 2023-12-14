using Newtonsoft.Json.Linq;
using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion
{
    /// <summary>
    /// Pipeline 2 launcher
    /// </summary>
    public class Pipeline2
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        /// <summary>
        /// Get the pipeline 2 root directory
        /// </summary>
        public static string InstallationPath
        {
            get
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                    + @"\daisy-pipeline";
            }
        }

        private JavaNativeInterface jni;

        // Classes
        private IntPtr SimpleAPIClass,
            CommandLineJobClass,
            JobStatusClass;

        #region Singleton initialisation

        private static Pipeline2 instance = null;

        // for thread safety
        private static readonly object padlock = new object();

        private static string getAsFileURI(string path)
        {
            return Uri.EscapeUriString(
                new Uri("file:/" + path.Replace("\\", "/")).ToString().Replace(":///", ":/")
            );
        }

        private Pipeline2()
        {
            // use jnet to execute the conversion on the inputPath
            // check for arch specific jre based on how the jre folders are created by the daisy/pipeline-assembly project (lite-bridge version)
            string arch = (IntPtr.Size * 8).ToString();
            string jrePath = Directory.Exists(InstallationPath + @"\jre" + arch)
                ? InstallationPath + @"\jre" + (IntPtr.Size * 8).ToString()
                : InstallationPath + @"\jre";
            string jvmDllPath = Path.Combine(jrePath, "bin", "server", "jvm.dll");

            Dictionary<string, string> SystemProps = new Dictionary<string, string>
            {
                {
                    "-Dorg.daisy.pipeline.properties",
                    "\"" + Path.Combine(InstallationPath, "etc", "pipeline.properties") + "\""
                },
                // Logback configuration file
                {
                    "-Dlogback.configurationFile",
                    getAsFileURI(Path.Combine(InstallationPath, "etc", "logback.xml"))
                },
                // Version number as returned by "alive" call
                { "-Dorg.daisy.pipeline.version", "" + PipelineVersion },
                // Workaround for encoding bugs on Windows
                { "-Dfile.encoding", "UTF8" },
                // to make ${org.daisy.pipeline.data}, ${org.daisy.pipeline.logdir} and ${org.daisy.pipeline.mode}
                // available in config-logback.xml and felix.properties
                // note that config-logback.xml is the only place where ${org.daisy.pipeline.mode} is used
                { "-Dorg.daisy.pipeline.data", AppDataFolder },
                { "-Dorg.daisy.pipeline.logdir", LogsFolder.Replace("\\", "/") },
                { "-Dorg.daisy.pipeline.mode", "cli" }
            };
            if (GlobaleSettings.AzureSpeechKey != "")
            {
                SystemProps["-Dorg.daisy.pipeline.tts.azure.key"] = GlobaleSettings.AzureSpeechKey;
                SystemProps["-Dorg.daisy.pipeline.tts.azure.region"] =
                    GlobaleSettings.AzureSpeechRegion;
            }

            List<string> JarPathes = ClassFolders.Aggregate(
                new List<string>(),
                (List<string> classPath, string path) =>
                {
                    return Directory.Exists(path)
                        ? classPath
                            .Concat(
                                Directory.EnumerateFiles(path, "*.jar", SearchOption.AllDirectories)
                            // .Select( fullPath => fullPath.Remove(0, InstallationPath.Length) ) // if needed to convert fullpath to installation relative path
                            )
                            .ToList()
                        : classPath;
                }
            );

            string ClassPath =
                JarPathes.Aggregate((acc, path) => acc + Path.PathSeparator + path)
                + Path.PathSeparator
                + Path.Combine(InstallationPath, "system", "simple-api");

            List<string> options = new List<string>();
            options = JavaOptions
                .Concat(
                    SystemProps.Aggregate(
                        new List<string>(),
                        (List<string> opts, KeyValuePair<string, string> opt) =>
                        {
                            opts.Add(opt.Key.ToString() + "=" + opt.Value.ToString());
                            return opts;
                        }
                    )
                )
                .ToList();
            options.Add("-Djava.class.path=" + ClassPath);
#if DEBUG // Add debugging capabilities on JVM
            options.Add("-Xdebug");
            options.Add("-XX:+CreateMinidumpOnCrash");
            // check if the jre/release file contains "jdk.jdwp.agent" in its modules list
            if (
                File.Exists(Path.Combine(jrePath, "release"))
                && File.ReadAllText(Path.Combine(jrePath, "release")).Contains("jdk.jdwp.agent")
            )
            {
                // activate debug capabilities if so
                options.Add("-agentlib:jdwp=transport=dt_socket,server=y,suspend=n,address=5005");
            }
#endif

            // Load a new JVM
            jni = new JavaNativeInterface(options, jvmDllPath, false);

            // Initialize runner in the JVM
            SimpleAPIClass = jni.GetJavaClass("SimpleAPI");
            CommandLineJobClass = jni.GetJavaClass("SimpleAPI$CommandLineJob");
            JobStatusClass = jni.GetJavaClass("org/daisy/pipeline/job/Job$Status");
            
            string systemOut = Path.Combine(LogsFolder, "sysOut.log");
            string systemErr = Path.Combine(LogsFolder, "sysErr.log");

            IntPtr JavaSystem = jni.GetJavaClass("java/lang/System");
            jni.CallVoidMethod(
                JavaSystem,
                IntPtr.Zero,
                "setOut",
                "(Ljava/io/PrintStream;)V",
                makePrintStream(systemOut)
            );
            jni.CallVoidMethod(
                JavaSystem,
                IntPtr.Zero,
                "setErr",
                "(Ljava/io/PrintStream;)V",
                makePrintStream(systemErr)
            );
        }

        private IntPtr makePrintStream(string filepath)
        {
            IntPtr PrintStream = jni.GetJavaClass("java/io/PrintStream");
            IntPtr BufferedOutputStream = jni.GetJavaClass("java/io/BufferedOutputStream");
            IntPtr FileOutputStream = jni.GetJavaClass("java/io/FileOutputStream");

            IntPtr FileOutputStreamObject = jni.NewObject(
                FileOutputStream,
                "(Ljava/lang/String;)V",
                filepath
            );
            IntPtr BufferedOutputStreamObject = jni.NewObject(
                BufferedOutputStream,
                "(Ljava/io/OutputStream;)V",
                FileOutputStreamObject
            );
            IntPtr PrintStreamObject = jni.NewObject(
                PrintStream,
                "(Ljava/io/OutputStream;)V",
                BufferedOutputStreamObject
            );

            return PrintStreamObject;
        }

        private void Dispose()
        {
            jni.Dispose();
            jni = null;
        }

        public static Pipeline2 Instance
        {
            get
            {
                lock (padlock)
                {
                    if (instance == null)
                    {
                        instance = new Pipeline2();
                    }
                    return instance;
                }
            }
        }

        public static void KillInstance()
        {
            lock (padlock)
            {
                if (instance != null)
                {
                    instance.Dispose();
                    instance = null;
                }
            }
        }
        #endregion


        #region JVM options (extracted from batch and bash launchers)

        private static string PipelineVersion = "1.14.15-SNAPSHOT";

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

        private static List<string> ClassFolders = new List<string>
        {
            Path.Combine(InstallationPath, "system", "common"),
            Path.Combine(InstallationPath, "system", "no-osgi"),
            Path.Combine(InstallationPath, "system", "simple-api")
        };

        private static List<string> JavaOptions = new List<string>
        {
            //"-server",
            "-Dcom.sun.management.jmxremote",
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
            "--add-exports=jdk.naming.rmi/com.sun.jndi.url.rmi=ALL-UNNAMED"
        };

        #endregion

        public delegate void PipelineOutputListener(string message);
        private event PipelineOutputListener OnPipelineOutputEvent;

        public void SetPipelineOutputListener(PipelineOutputListener onPipelineOutput)
        {
            this.OnPipelineOutputEvent = onPipelineOutput;
        }

        public PipelineOutputListener OnPipelineOutput => OnPipelineOutputEvent;

        public delegate void PipelineErrorListener(string message);
        private event PipelineErrorListener OnPipelineErrorEvent;

        public void SetPipelineErrorListener(PipelineErrorListener onPipelineError)
        {
            this.OnPipelineErrorEvent = onPipelineError;
        }

        public PipelineErrorListener OnPipelineError => this.OnPipelineErrorEvent;

        public IntPtr Start(string scriptName, Dictionary<string, object> options = null)
        {
            try
            {
                IntPtr hashMap = jni.NewJavaWrapperObject(
                    options != null ? options : new Dictionary<string, object>()
                );

                IntPtr job = jni.CallMethod<IntPtr>(
                    SimpleAPIClass,
                    IntPtr.Zero,
                    "startJob",
                    "(Ljava/lang/String;Ljava/util/Map;)LSimpleAPI$CommandLineJob;",
                    scriptName,
                    jni.CallMethod<IntPtr>(
                        SimpleAPIClass,
                        IntPtr.Zero,
                        "stringifyOptions",
                        "(Ljava/util/Map;)Ljava/util/Map;",
                        hashMap
                    )
                );

                return job;
            }
            catch (Exception e)
            {
                if (OnPipelineError != null)
                {
                    string errorMessage = e.Message;

                    while (e.InnerException != null)
                    {
                        e = e.InnerException;
                        errorMessage += "\r\n - Thrown by " + e.Message;
                    }
                    errorMessage += "\r\n\r\n" + e.StackTrace;
                    OnPipelineError(errorMessage);
                }
                throw;
                //return IntPtr.Zero;
            }
        }

        /// <summary>
        /// Possible status for a pipeline job
        /// </summary>
        public enum JobStatus
        {
            IDLE,
            RUNNING,
            SUCCESS,
            ERROR,
            FAIL,
            UNKNOWN
        }

        /// <summary>
        /// Retrieve the current status of a pipeline job
        /// </summary>
        /// <param name="jobObject"></param>
        /// <returns></returns>
        public JobStatus getStatus(IntPtr jobObject)
        {
            if (jobObject == IntPtr.Zero)
            {
                throw new Exception("Status requested on non-existing job");
            }
            string status = jni.CallMethod<string>(
                JobStatusClass,
                jni.CallMethod<IntPtr>(
                    CommandLineJobClass,
                    jobObject,
                    "getStatus",
                    "()Lorg/daisy/pipeline/job/Job$Status;"
                ),
                "toString",
                "()Ljava/lang/String;"
            );
            switch (status)
            {
                case "IDLE":
                    return JobStatus.IDLE;
                case "RUNNING":
                    return JobStatus.RUNNING;
                case "SUCCESS":
                    return JobStatus.SUCCESS;
                case "ERROR":
                    return JobStatus.ERROR;
                case "FAIL":
                    return JobStatus.FAIL;
                default:
                    return JobStatus.UNKNOWN;
            }
        }

        public List<string> getNewMessages(IntPtr commandLinejob)
        {
            if (commandLinejob == IntPtr.Zero) {
                throw new Exception("getNewMessages requested on non-existing job");
            }
            List<string> messages = new List<string>();

            IntPtr errorsList = jni.CallMethod<IntPtr>(
                CommandLineJobClass,
                commandLinejob,
                "getNewMessages",
                "()Ljava/util/List;"
            );

            JavaIterable<IntPtr> messagesIterable = new JavaIterable<IntPtr>(jni, errorsList);
            IntPtr messageClass = IntPtr.Zero;
            foreach (IntPtr messageObject in messagesIterable) {
                if (messageClass == IntPtr.Zero) {
                    messageClass = jni.JNIEnvironment.GetObjectClass(messageObject);
                }
                string message = jni.CallMethod<string>(
                    messageClass,
                    messageObject,
                    "getText",
                    "()Ljava/lang/String;"
                );
                messages.Add(message);
            }
            return messages;
        }

        public List<string> getErros(IntPtr commandLinejob)
        {
            if (commandLinejob == IntPtr.Zero) {
                throw new Exception("getErros requested on non-existing job");
            }
            List<string> messages = new List<string>();

            IntPtr errorsList = jni.CallMethod<IntPtr>(
                CommandLineJobClass,
                commandLinejob,
                "getErrors",
                "()Ljava/util/List;"
            );

            JavaIterable<IntPtr> messagesIterable = new JavaIterable<IntPtr>(jni, errorsList);
            IntPtr messageClass = IntPtr.Zero;
            foreach (IntPtr messageObject in messagesIterable) {
                if (messageClass == IntPtr.Zero) {
                    messageClass = jni.JNIEnvironment.GetObjectClass(messageObject);
                }
                string message = jni.CallMethod<string>(
                    messageClass,
                    messageObject,
                    "getText",
                    "()Ljava/lang/String;"
                );
                messages.Add(message);
            }
            return messages;
        }
    }
}
