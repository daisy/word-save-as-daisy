using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace ConversionApp
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

        public delegate void JobStandardOutput(string message);

        public delegate void JobProgressOutput(int current, int total = int.MaxValue);

        public delegate void JobErrorOutput(string message);


        public static async Task<int> StartJob(
            string scriptOrCommand,
            Dictionary<string, string> options = null,
            CancellationToken cancellationToken = default,
            JobStandardOutput info = null,
            JobProgressOutput progress = null,
            JobErrorOutput error = null
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
                        
                        JavaNativeInterface jni = startJNIBridge();
                        IntPtr SimpleAPISPIClass = jni.GetJavaClass("SimpleAPI_SPI");
                        IntPtr SimpleAPIClass = jni.GetJavaClass("SimpleAPI");
                        IntPtr CommandLineJobClass = jni.GetJavaClass("SimpleAPI$CommandLineJob");
                        IntPtr JobStatusClass = jni.GetJavaClass("org/daisy/pipeline/job/Job$Status");
                        IntPtr SimpleAPIInstance = jni.NewObject(SimpleAPISPIClass);

                        switch (scriptOrCommand)
                        {
                            case "datatypes":
                                break;
                            case "properties":
                                break;
                            case "scripts":
                                break;
                            default:
                                // The command is a script
                                break;
                        }


                        IntPtr currentJob = IntPtr.Zero;

                        try
                        {
                            Dictionary<string, List<string>> stringifiedOptions =
                                new Dictionary<string, List<string>>();
                            foreach (KeyValuePair<string, string> option in options)
                            {
                                if (option.Value != null)
                                {
                                    // Note : the hashmap expects list of strings as options values
                                    stringifiedOptions[option.Key] = new List<string>()
                                {
                                    StringifyOption(jni, option.Value),
                                    //option.Value.ToString() // possible only after referenced pipeline framework is updated to enable a more relaxed value check (https://github.com/daisy/pipeline-framework/commit/46dc1aeb6918da24640d327c7f6cd2b0c44b1dd5)
                                };
                                }

                            }
                            IntPtr hashMap = jni.NewJavaWrapperObject(stringifiedOptions);
                            info?.Invoke($"Launching the conversion job...");
                            //Dispatcher.Invoke(() =>
                            //{
                            //    string message = $"Launching the conversion job...";
                            //    ConversionProgressText.Text = message;
                            //    Console.WriteLine(message);
                            //    LogTextBox.AppendText(message + "\r\n");
                            //    LogTextBox.ScrollToEnd();
                            //});

                            currentJob = jni.CallMethod<IntPtr>(
                                SimpleAPIClass,
                                SimpleAPIInstance,
                                "startJob",
                                "(Ljava/lang/String;Ljava/util/Map;)LSimpleAPI$CommandLineJob;",
                                scriptOrCommand,
                                hashMap
                            );
                        }
                        catch (Exception e)
                        {
                            Exception scriptError = new Exception(
                                $"Script {scriptOrCommand} raised an error",
                                e
                            );
                            throw scriptError;
                            //return IntPtr.Zero;
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
                                    //Dispatcher.Invoke(() =>
                                    //{
                                    //    ConversionProgressText.Text = message;
                                    //    LogTextBox.AppendText("DP2 > " + message + "\r\n");
                                    //    LogTextBox.ScrollToEnd();
                                    //});
                                    //Console.WriteLine("DP2 > " + message);
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
                                        string errorMessage =
                                            " DP2 > "
                                            + scriptOrCommand
                                            + " conversion job has finished in error :\r\n"
                                            + string.Join("\r\n", errors);
                                       
                                        throw new JobException(errorMessage);
                                    case JobStatus.Fail:
                                        checkStatus = false;
                                        errors = getErros(jni, currentJob, CommandLineJobClass);
                                        string failedMessage =
                                            " DP2 > "
                                            + scriptOrCommand
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
                                "DP2 > An unknown error occured while launching the script "
                                    + scriptOrCommand
                                    + " with the parameters "
                                    + options.Aggregate(
                                        "",
                                        (result, keyvalue) =>
                                            result + keyvalue.Key + "=" + keyvalue.Value.ToString() + "\r\n"
                                    )
                            // + "Please try to run the conversion in DAISY Pipeline 2 application directly for more informations on the issue."
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
            catch (AggregateException aggEx)
            {
                error?.Invoke("An error occured during the conversion : ");
                foreach (var ex in aggEx.InnerExceptions)
                {
                    error?.Invoke(ex.Message);
                    error?.Invoke(ex.StackTrace);
                }
                return 2;
            }
            catch (JobException jobEx)
            {
                error?.Invoke(jobEx.Message);
                error?.Invoke(jobEx.StackTrace);
                return 2;
            }
            catch (Exception ex)
            {
                error?.Invoke(ex.Message);
                error?.Invoke(ex.StackTrace);
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

        public static JavaNativeInterface startJNIBridge(IConversionEventsHandler events = null)
        {
            List<string> ClassFolders = new List<string>
            {
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "common"),
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "no-osgi"),
                Path.Combine(ConverterHelper.EmbeddedEnginePath, "system", "simple-api")
            };

            List<string> JavaOptions = new List<string>
            {
                //"-server",
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
