using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;
using static org.daisy.jniwrapper.JNITaskRunner;

namespace org.daisy.jniwrapper
{
    internal class JNITaskRunner
    {
        private static string StringifyOption(JavaNativeInterface jni, object obj)
        {
            IntPtr javaClass;
            string value = obj.ToString();
            if (obj is string)
            {
                javaClass = jni.GetJavaClass("java/lang/String");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.NewObject(javaClass, "(Ljava/lang/String;)V", (string)obj),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is long)
            {
                javaClass = jni.GetJavaClass("java/lang/Long");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(J)Ljava/lang/Long;",
                        new object[1] { (long)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is int)
            {
                javaClass = jni.GetJavaClass("java/lang/Integer");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(I)Ljava/lang/Integer;",
                        new object[1] { (int)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is short)
            {
                javaClass = jni.GetJavaClass("java/lang/Short");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(S)Ljava/lang/Short;",
                        new object[1] { (short)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is bool)
            {
                javaClass = jni.GetJavaClass("java/lang/Boolean");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(Z)Ljava/lang/Boolean;",
                        new object[1] { (bool)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is float)
            {
                javaClass = jni.GetJavaClass("java/lang/Float");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(F)Ljava/lang/Float;",
                        new object[1] { (float)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            else if (obj is double)
            {
                javaClass = jni.GetJavaClass("java/lang/Double");
                value = jni.CallMethod<string>(
                    javaClass,
                    jni.CallMethod<IntPtr>(
                        javaClass,
                        IntPtr.Zero,
                        "valueOf",
                        "(D)Ljava/lang/Double;",
                        new object[1] { (double)obj }
                    ),
                    "toString",
                    "()Ljava/lang/String;"
                );
            }
            return value;
        }

        /// <summary>
        /// Retrieve the current status of a pipeline job
        /// </summary>
        /// <param name="jobObject"></param>
        /// <returns></returns>
        private static JobStatus getStatus(
            JavaNativeInterface jni,
            IntPtr jobObject,
            IntPtr CommandLineJobClass,
            IntPtr JobStatusClass
        )
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
                case "RUNNING":
                    return JobStatus.Running;
                case "SUCCESS":
                    return JobStatus.Success;
                case "ERROR":
                    return JobStatus.Error;
                case "FAIL":
                    return JobStatus.Fail;
                case "IDLE":
                default:
                    return JobStatus.Idle;
            }
        }

        public static List<string> getNewMessages(
            JavaNativeInterface jni,
            IntPtr commandLinejob,
            IntPtr CommandLineJobClass
        )
        {
            if (commandLinejob == IntPtr.Zero)
            {
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
            foreach (IntPtr messageObject in messagesIterable)
            {
                if (messageClass == IntPtr.Zero)
                {
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

        private static List<string> getErros(
            JavaNativeInterface jni,
            IntPtr commandLinejob,
            IntPtr CommandLineJobClass
        )
        {
            if (commandLinejob == IntPtr.Zero)
            {
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
            foreach (IntPtr messageObject in messagesIterable)
            {
                if (messageClass == IntPtr.Zero)
                {
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

        public delegate void StandardOutput(string message);

        public delegate void ProgressOutput(int current, int total = int.MaxValue);

        public delegate void ErrorOutput(string message);

        public enum Command
        {
            None = 0,
            Descriptors = 1,
            Scripts = 2,
            ScriptDetails = 3,
            Datatypes = 4,
            DatatypeDetails = 5,
            SettableProperties = 6,
        }

        public static async Task<int> ExecuteCommand(
            Command command,
            Dictionary<string, string> options = null,
            CancellationToken cancellationToken = default,
            StandardOutput info = null,
            ProgressOutput progress = null,
            ErrorOutput error = null,
            Dictionary<string, string> properties = null
        ) {
            try
            {
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        info?.Invoke($"Starting JVM and conversion pipeline...");

                        JavaNativeInterface jni = startJNIBridge(properties: properties);
                        IntPtr SimpleAPISPIClass = jni.GetJavaClass("SimpleAPI_SPI");
                        IntPtr SimpleAPIClass = jni.GetJavaClass("SimpleAPI");
                        IntPtr SimpleAPIInstance = jni.NewObject(SimpleAPISPIClass);

                        string outputDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        string outputFile = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        
                        if (options.ContainsKey("output"))
                        {
                            outputDirectory = options["output"];
                            outputFile = options["output"];
                            // Output selected is a directory, we create a file name based on the command inside it
                            if (Directory.Exists(outputFile))
                            {
                                outputFile = Path.Combine(outputFile, $"{command.ToString().ToLower()}.xml");
                            } else
                            {
                                outputDirectory = Path.GetDirectoryName(outputFile);
                            }
                        }

                        string scriptsDescriptors;
                        string datatypesDescriptors;
                        string result;
                        switch (command)
                        {
                            case Command.Descriptors:
                                info?.Invoke($"Retrieving all descriptors...");
                                scriptsDescriptors = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getScripts",
                                    "(Z)Ljava/lang/String;",
                                    true
                                );
                                File.WriteAllText(Path.Combine(outputDirectory, "scripts.xml"), scriptsDescriptors);
                                datatypesDescriptors = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getDatatypes",
                                    "()Ljava/lang/String;"
                                );
                                File.WriteAllText(Path.Combine(outputDirectory, "datatypes.xml"), datatypesDescriptors);
                                break;
                            case Command.Scripts:
                                info?.Invoke($"Retrieving scripts descriptors...");
                                result = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getScripts",
                                    "(Z)Ljava/lang/String;",
                                    true
                                );
                                File.WriteAllText(outputFile, result);
                                break;
                            case Command.ScriptDetails:
                                if (!options.ContainsKey("id"))
                                {
                                    throw new Exception("The 'id' option is required for the ScriptDetails command");
                                }
                                string scriptId = options["id"];
                                info?.Invoke($"Retrieving details for script {scriptId}...");
                                result = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getScriptDetails",
                                    "(Ljava/lang/String;)Ljava/lang/String;",
                                    scriptId
                                );
                                File.WriteAllText(outputFile, result);
                                break;
                            case Command.Datatypes:
                                info?.Invoke($"Retrieving datatypes descriptors...");
                                result = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getDatatypes",
                                    "()Ljava/lang/String;"
                                );
                                File.WriteAllText(outputFile, result);
                                break;
                            case Command.DatatypeDetails:
                                if (!options.ContainsKey("id"))
                                {
                                    throw new Exception("The 'id' option is required for the DatatypeDetails command");
                                }
                                string datatypeId = options["id"];
                                info?.Invoke($"Retrieving details for datatype {datatypeId}...");
                                result = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getDatatypeDetails",
                                    "(Ljava/lang/String;)Ljava/lang/String;",
                                    datatypeId
                                );
                                File.WriteAllText(outputFile, result);
                                break;
                            case Command.SettableProperties:
                                info?.Invoke($"Retrieving settable properties descriptors...");
                                result = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getSettableProperties",
                                    "()Ljava/lang/String;"
                                );
                                File.WriteAllText(outputFile, result);
                                break;
                            default:
                                throw new Exception("Unknown command : " + command.ToString());
                        }

                    }
                    catch (Exception)
                    {
                        throw;
                    }
                });
            }
            catch (OperationCanceledException)
            {
                info?.Invoke("Export was cancelled.");
                return 1;
            }
            catch (AggregateException aggEx)
            {
                StringBuilder errorLog = new StringBuilder($"An error occured during command {command} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                
                error?.Invoke($"An error occured during command {command} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                foreach (var ex in aggEx.InnerExceptions)
                {
                    errorLog.AppendLine(ex.Message);
                    errorLog.AppendLine(ex.StackTrace);
                }
                string path = Path.Combine(JNILogsPath, $"{command}-error-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                File.WriteAllText(path, errorLog.ToString());
                error?.Invoke($"For more details, consult the error log at {path}");
                return 2;
            }
            catch (Exception ex)
            {
                StringBuilder errorLog = new StringBuilder($"An error occured during command {command} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");

                error?.Invoke($"An error occured during command {command} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                while (ex != null)
                {
                    error?.Invoke($"{ex.Message}");
                    errorLog.AppendLine(ex.Message);
                    errorLog.AppendLine(ex.StackTrace);
                    ex = ex.InnerException;
                }
                string path = Path.Combine(JNILogsPath, $"{command}-error-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                File.WriteAllText(path, errorLog.ToString());
                error?.Invoke($"For more details, consult the error log at {path}");

                return 3;
            }

            return 0;
        }

        /// <summary>
        /// Launch a conversion using the SimpleAPI class through JNI
        /// The method will start the JVM, launch the conversion and monitor its status until completion, error or cancellation.
        /// </summary>
        /// <param name="script"></param>
        /// <param name="options"></param>
        /// <param name="cancellationToken"></param>
        /// <param name="info"></param>
        /// <param name="progress"></param>
        /// <param name="error"></param>
        /// <returns></returns>
        public static async Task<int> LaunchConversion(
            string script,
            Dictionary<string, string> options = null,
            CancellationToken cancellationToken = default,
            StandardOutput info = null,
            ProgressOutput progress = null,
            ErrorOutput error = null,
            Dictionary<string, string> properties = null
        )
        {
            try
            {
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        info?.Invoke($"Starting JVM and conversion pipeline...");
                        
                        JavaNativeInterface jni = startJNIBridge(properties: properties);
                        IntPtr SimpleAPISPIClass = jni.GetJavaClass("SimpleAPI_SPI");
                        IntPtr SimpleAPIClass = jni.GetJavaClass("SimpleAPI");
                        IntPtr CommandLineJobClass = jni.GetJavaClass("SimpleAPI$CommandLineJob");
                        IntPtr JobStatusClass = jni.GetJavaClass("org/daisy/pipeline/job/Job$Status");
                        IntPtr SimpleAPIInstance = jni.NewObject(SimpleAPISPIClass);

                        // Retrieve the script definition
                        string definition = jni.CallMethod<string>(
                                    SimpleAPIClass,
                                    SimpleAPIInstance,
                                    "getScriptDetails",
                                    "(Ljava/lang/String;)Ljava/lang/String;",
                                    script
                                );
                        XmlDocument definitionDoc = new XmlDocument();
                        definitionDoc.LoadXml(definition);
                        ScriptDefinition scriptDefinition = ScriptDefinition.FromXml(definitionDoc.DocumentElement);
                       
                        
                        IntPtr currentJob = IntPtr.Zero;
                        List<string> missingRequiredItems = scriptDefinition.Options
                                    .Concat<ScriptItemBase>(scriptDefinition.Inputs)
                                    .Concat<ScriptItemBase>(scriptDefinition.Outputs)
                                    .Where(ScriptItemBase => ScriptItemBase.Required == true)
                                    .Select(ScriptItemBase=> ScriptItemBase.Name)
                                    .Where(o => !options.ContainsKey(o) || options[o] == null)
                                    .ToList();
                        if (missingRequiredItems.Count > 0)
                        {
                            throw new JobException($"The following required fields are missing or empty for the script {script} : {string.Join(", ", missingRequiredItems)}");
                        }
                        try
                        {
                            Dictionary<string, List<string>> stringifiedOptions =
                                new Dictionary<string, List<string>>();

                            foreach (KeyValuePair<string, string> option in options)
                            {
                                ScriptItemBase item = scriptDefinition.Options
                                    .Concat<ScriptItemBase>(scriptDefinition.Inputs)
                                    .Concat<ScriptItemBase>(scriptDefinition.Outputs)
                                    .FirstOrDefault(opt => opt.Name == option.Key);
                                if (item == null)
                                {
                                    info?.Invoke(
                                        $"Warning : the field {option.Key} is not defined for the script {script}, it will be ignored"
                                    );
                                }
                                else if (option.Value == null)
                                {
                                    info?.Invoke(
                                        $"Warning : the field {option.Key} is set to null, it will be ignored"
                                    );
                                } else {
                                    // Note : the hashmap expects list of strings as options values
                                    stringifiedOptions[option.Key] = new List<string>()
                                    {
                                        StringifyOption(jni, option.Value),
                                    };
                                }
                            }
                            IntPtr hashMap = jni.NewJavaWrapperObject(stringifiedOptions);
                            info?.Invoke($"Launching the conversion job...");

                            currentJob = jni.CallMethod<IntPtr>(
                                SimpleAPIClass,
                                SimpleAPIInstance,
                                "startJob",
                                "(Ljava/lang/String;Ljava/util/Map;)LSimpleAPI$CommandLineJob;",
                                script,
                                hashMap
                            );
                        }
                        catch (Exception e)
                        {
                            Exception scriptError = new Exception(
                                $"Could not start a job for {script} with the provided options.",
                                e
                            );
                            throw scriptError;
                        }

                        if (currentJob != IntPtr.Zero)
                        {
                            bool checkStatus = true;
                            List<string> errors;
                            int _progress = 0;
                            while (checkStatus && !cancellationToken.IsCancellationRequested)
                            {
                                foreach (
                                    string message in getNewMessages(jni, currentJob, CommandLineJobClass)
                                )
                                {
                                    info?.Invoke(message);
                                }
                                progress?.Invoke(_progress++);
                                switch (getStatus(jni, currentJob, CommandLineJobClass, JobStatusClass))
                                {
                                    case JobStatus.Idle:
                                        break;
                                    case JobStatus.Running:
                                        break;
                                    case JobStatus.Success:
                                        checkStatus = false;
                                        info?.Invoke("Conversion completed successfully.");
                                        progress?.Invoke(100,100);
                                        break;
                                    case JobStatus.Error:
                                        errors = getErros(jni, currentJob, CommandLineJobClass);
                                        string errorMessage = script
                                            + " conversion job has finished in error :\r\n"
                                            + string.Join("\r\n", errors);
                                       
                                        throw new JobException(errorMessage);
                                    case JobStatus.Fail:
                                        checkStatus = false;
                                        errors = getErros(jni, currentJob, CommandLineJobClass);
                                        string failedMessage = script
                                            + " conversion job failed :\r\n"
                                            + string.Join("\r\n", errors);
                                       
                                        throw new JobException(failedMessage);
                                    default:
                                        break;
                                }
                                cancellationToken.ThrowIfCancellationRequested();
                            }
                        }
                        else
                        {
                            throw new Exception(
                                "An unknown error occured while launching the script "
                                    + script
                                    + " with the parameters "
                                    + options.Aggregate(
                                        "",
                                        (result, keyvalue) =>
                                            result + keyvalue.Key + "=" + keyvalue.Value.ToString() + "\r\n"
                                    )
                            );
                        }
                    }
                    catch (JobException)
                    {
                        throw;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    finally
                    {
                    }
                });
                return 0;
            }
            catch (OperationCanceledException)
            {
                info?.Invoke("Conversion was cancelled.");
                return 1;
            }
            catch (JobException ex)
            {
                StringBuilder errorLog = new StringBuilder($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");

                error?.Invoke($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                error?.Invoke($"{ex.Message}");
                errorLog.AppendLine(ex.Message);
                errorLog.AppendLine(ex.StackTrace);
                
                string path = Path.Combine(JNILogsPath, $"{script}-joberror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                File.WriteAllText(path, errorLog.ToString());
                error?.Invoke($"For more details, consult the error log at {path}");
                MessageBox.Show(
                    $"An error occured during the conversion : \r\n {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );

                return 2;
            }
            catch (AggregateException aggEx)
            {
                StringBuilder errorLog = new StringBuilder($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");

                error?.Invoke($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                foreach (var ex in aggEx.InnerExceptions)
                {
                    errorLog.AppendLine(ex.Message);
                    errorLog.AppendLine(ex.StackTrace);
                }
                string path = Path.Combine(AppDataFolder, $"{script}-criticalerror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                File.WriteAllText(path, errorLog.ToString());
                error?.Invoke($"For more details, consult the error log at {path}");
                return 2;
            }
            catch (Exception ex)
            {
                StringBuilder errorLog = new StringBuilder($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");

                error?.Invoke($"An error occured running {script} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}");
                while (ex != null)
                {
                    error?.Invoke($"{ex.Message}");
                    errorLog.AppendLine(ex.Message);
                    errorLog.AppendLine(ex.StackTrace);
                    ex = ex.InnerException;
                }
                string path = Path.Combine(JNILogsPath, $"export-error-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                File.WriteAllText(path, errorLog.ToString());
                error?.Invoke($"For more details, consult the error log at {path}");

                if (MessageBox.Show(
                    $"An error occured during the conversion. Do you want to consult the error log at {path}?",
                    "Error",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Error
                ) == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }

                return 3;
            }
        }

        private static string getAsFileURI(string path)
        {
            return Uri.EscapeUriString(
                new Uri("file:/" + path.Replace("\\", "/")).ToString().Replace(":///", ":/")
            );
        }
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

        public static JavaNativeInterface startJNIBridge(IConversionEventsHandler events = null, Dictionary<string, string> properties = null)
        {
            List<string> ClassFolders = new List<string>
            {
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "common"),
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "no-osgi"),
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api")
            };

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
            string jvmDllPath = Path.Combine(jrePath, "bin", "server", "jvm.dll");

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
                { "-Dorg.daisy.pipeline.data", AppDataFolder },
                { "-Dorg.daisy.pipeline.logdir", LogsFolder.Replace("\\", "/") },
                { "-Dorg.daisy.pipeline.mode", "cli" }
            };

            if (properties != null && properties.Count > 0)
            {
                
                foreach (var prop in properties)
                {
                    if (SystemProps.ContainsKey(prop.Key))
                    {
                        SystemProps[prop.Key] = prop.Value;
                    }
                    else
                    {
                        SystemProps.Add(prop.Key, prop.Value);
                    }
                }
            }

            // Load properties in the JNI runner
            //if (ConverterHelper.EmbeddedEngineProperties != null && File.Exists(ConverterHelper.EmbeddedEngineProperties))
            //{
            //    try
            //    {
            //        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            //        xmlDoc.Load(ConverterHelper.EmbeddedEngineProperties);
            //        var element = xmlDoc.GetElementsByTagName("properties");
            //        var props = EngineProperty.ListFromXml((System.Xml.XmlElement)element.Item(0));
            //        foreach (var prop in props)
            //        {
            //            SystemProps.Add("-D" + prop.Name, prop.Value);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        // No loading of the settable properties file if it is not a well-formed XML, but log the error and continue with the default properties
            //    }
            //}

            List<string> JarPathes = ClassFolders.Aggregate(
                new List<string>(),
                (classPath, path) =>
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
                + Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api");

            List<string> options = new List<string>();
            options = JavaOptions
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
            options.Add("-Djava.class.path=" + ClassPath);

            return new JavaNativeInterface(options, jvmDllPath, false);

        }

    }
}
