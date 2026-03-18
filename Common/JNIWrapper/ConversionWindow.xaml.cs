
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ConversionApp
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class ConversionWindow : Window
    {
        
        private CancellationTokenSource cancellationTokenSource;
        private Dictionary<string, string> parsedOptions;

        public ConversionWindow(string scriptName, Dictionary<string, string> parsedOptions)
        {

            InitializeComponent();
            //string[] args = Environment.GetCommandLineArgs();
            //string scriptName;
            //parsedOptions = ParseCommandLineArgs(args, out scriptName);

            this.Title = "DAISY Pipeline 2 - " + scriptName + " conversion in progress...";
            ConversionProgressText.Text = "Running " + scriptName + " conversion...";
            LogTextBox.AppendText("Starting conversion with the following options:\r\n");
            foreach (KeyValuePair<string, string> option in parsedOptions)
            {
                LogTextBox.AppendText($"--{option.Key} \"{option.Value}\"\r\n");
            }
            LogTextBox.ScrollToEnd();
            cancellationTokenSource = new CancellationTokenSource();

            JNITaskRunner.StartJob(scriptName, parsedOptions, cancellationTokenSource.Token,
                info: (m =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ConversionProgressText.Text = m;
                            LogTextBox.AppendText(m + "\r\n");
                            LogTextBox.ScrollToEnd();
                        });
                        Console.WriteLine("DP2 > " + m);
                    }),
                error: (m =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            LogTextBox.AppendText("DP2 > " + m + "\r\n");
                            LogTextBox.ScrollToEnd();
                        });
                        Console.WriteLine("DP2 > " + m);
                    }),
                progress: ((p,t) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Maximum = t;
                            ProgressBar.Value = p;
                        });
                    })
            ).ContinueWith(t =>
            {
                Dispatcher.Invoke(() =>
                {
                    Application.Current.Shutdown(t.Result);
                });
               
            });

            //this.Start(scriptName, parsedOptions, cancellationTokenSource.Token);
        }
        
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            cancellationTokenSource.Cancel();
            LogTextBox.AppendText("Conversion cancelled by user.\r\n");
            LogTextBox.ScrollToEnd();
            Console.WriteLine("Conversion cancelled by user.");
            this.Close();
            Application.Current.Shutdown(1);
        }

        /// <summary>
        /// Convert a C# raw-type object into a string representation of the corresponding Java object
        /// </summary>
        /// <param name="jni">JNI connection</param>
        /// <param name="obj">Object to convert to java-like string</param>
        /// <returns></returns>
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

        public List<string> getNewMessages(
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



        private async void Start(string scriptOrCommand, Dictionary<string, string> options = null, CancellationToken cancellationToken = default)
        {
            try
            {
                await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        Dispatcher.Invoke(() =>
                        {
                            string message = $"Starting JVM and conversion pipeline...";
                            ConversionProgressText.Text = message;
                            Console.WriteLine(message);
                            LogTextBox.AppendText(message + "\r\n");
                            LogTextBox.ScrollToEnd();
                        });
                        
                        JavaNativeInterface jni = Engine.startJNIBridge();
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
                            Dispatcher.Invoke(() =>
                            {
                                string message = $"Launching the conversion job...";
                                ConversionProgressText.Text = message;
                                Console.WriteLine(message);
                                LogTextBox.AppendText(message + "\r\n");
                                LogTextBox.ScrollToEnd();
                            });

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
                            int progress = 0;
                            while (checkStatus && !cancellationToken.IsCancellationRequested)
                            {
                                foreach (
                                    string message in getNewMessages(jni, currentJob, CommandLineJobClass)
                                )
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        ConversionProgressText.Text = message;
                                        LogTextBox.AppendText("DP2 > " + message + "\r\n");
                                        LogTextBox.ScrollToEnd();
                                    });
                                    Console.WriteLine("DP2 > " + message);
                                }
                                Dispatcher.Invoke(() =>
                                {
                                    ProgressBar.Value = progress;
                                });
                                switch (getStatus(jni, currentJob, CommandLineJobClass, JobStatusClass))
                                {
                                    case JobStatus.Idle:
                                        break;
                                    case JobStatus.Running:
                                        break;
                                    case JobStatus.Success:
                                        checkStatus = false;
                                        Dispatcher.Invoke(() =>
                                        {
                                            LogTextBox.AppendText("Conversion completed successfully.\r\n");
                                            LogTextBox.ScrollToEnd();
                                        });
                                        Console.WriteLine("Conversion completed successfully.");
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
            }
            catch (OperationCanceledException)
            {
                Dispatcher.Invoke(() =>
                {
                    LogTextBox.AppendText("Conversion was cancelled.\r\n");
                    LogTextBox.ScrollToEnd();
                });
                Console.WriteLine("Conversion was cancelled.");
                Application.Current.Shutdown(1);
            }
            catch (AggregateException aggEx)
            {
                foreach (var ex in aggEx.InnerExceptions)
                {
                    Console.Error.WriteLine("Error: " + ex.Message);
                    Console.Error.WriteLine(ex.StackTrace);
                }
                Application.Current.Shutdown(2);
            }
            catch (JobException jobEx)
            {
                //MessageBox.Show(jobEx.Message);
                Console.Error.WriteLine("Error: " + jobEx.Message);
                Console.Error.WriteLine(jobEx.StackTrace);
                Application.Current.Shutdown(2);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("An error occured during the conversion : " + ex.Message);
                Console.Error.WriteLine("Error: " + ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Application.Current.Shutdown(3);
            }
            Application.Current.Shutdown(0);
        }
    }
}
