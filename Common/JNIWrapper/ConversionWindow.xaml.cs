using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Path = System.IO.Path;

namespace org.daisy.jniwrapper
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class ConversionWindow : Window
    {
        
        private CancellationTokenSource cancellationTokenSource;

        public ConversionWindow(string scriptName, Dictionary<string, string> options, Dictionary<string, string> properties = null)
        {

            InitializeComponent();
            //string[] args = Environment.GetCommandLineArgs();
            //string scriptName;
            //parsedOptions = ParseCommandLineArgs(args, out scriptName);

            this.Title = "DAISY Pipeline 2 - " + scriptName + " conversion in progress...";
            ConversionProgressText.Text = "Running " + scriptName + " conversion...";
            LogTextBox.AppendText("Starting conversion with the following options:\r\n");
            foreach (KeyValuePair<string, string> option in options)
            {
                LogTextBox.AppendText($"--{option.Key} \"{option.Value}\"\r\n");
            }
            LogTextBox.ScrollToEnd();
            cancellationTokenSource = new CancellationTokenSource();
            Task.Run(async () =>
            {
                try
                {
                    await JNITaskRunner.RunConversion(scriptName, options, cancellationTokenSource.Token,
                    info: (m =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ConversionProgressText.Text = m;
                            LogTextBox.AppendText(m + "\r\n");
                            LogTextBox.ScrollToEnd();
                        });
                        Console.WriteLine(m);
                    }),
                    error: (m =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            LogTextBox.AppendText(m + "\r\n");
                            LogTextBox.ScrollToEnd();
                        });
                        Console.WriteLine(m);
                    }),
                    progress: ((p, t) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Maximum = t;
                            ProgressBar.Value = p;
                            ConversionProgressPercentage.Text = $"{(100.0 * ProgressBar.Value / ProgressBar.Maximum):0.00}%";
                        });
                    }),
                    properties: properties);
                }
                catch (Exception iex)
                {
                    string path = System.IO.Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-unknowerror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                    StringBuilder errorLog = new StringBuilder($"An error occured running {scriptName} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}\r\n");
                    StringBuilder errorMessage = new StringBuilder();     
                    var ex = iex;
                    int errorCode = 1;
                    while (ex != null)
                    {
                        if(!ex.Message.StartsWith("file:/")) // Avoid the detailed log
                        {
                            errorMessage.AppendLine(ex.Message);
                            errorLog.AppendLine(ex.Message);
                            errorLog.AppendLine(ex.StackTrace);
                        }
                        ex = ex.InnerException;
                    }
                    if (iex is OperationCanceledException)
                    {
                        Console.WriteLine("Conversion was cancelled.");

                    }
                    else if (iex is JobException jex)
                    {
                        path = Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-joberror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                        
                        if (iex.Message.StartsWith("file:/"))
                        {
                            try
                            {
                                string newPath = new Uri(iex.Message).LocalPath;
                                errorLog.AppendLine($"Please send the log at the following path to the DAISY Pipeline team for additional troubleshooting :");
                                errorLog.AppendLine($"{newPath}");
                                File.WriteAllText(path, errorLog.ToString());
                                path = newPath;
                            } catch (UriFormatException)
                            {
                                // If it fails, keep the original path where we logged the error
                            }
                        } else
                        {
                            File.WriteAllText(path, errorLog.ToString());
                        }
                        errorCode = 2;
                    }
                    else
                    {
                        path = Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-criticalerror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                        File.WriteAllText(path, errorLog.ToString());
                        errorCode = 3;
                    }
                    if (errorCode == 1)
                    {
                        return;
                    }
                    else
                    {
                        Dispatcher.Invoke(() =>
                        {
                            if (MessageBox.Show(
                                    $"The following error occured during the conversion:" +
                                    $"\r\n{errorMessage}" +
                                    $"\r\nDo you want to open the conversion log for more details?" +
                                    $"\r\n(Please send those logs after review to the DAISY Pipeline team if you need additional troubleshooting.)",
                                    "Error during conversion",
                                    MessageBoxButton.YesNo,
                                    MessageBoxImage.Error
                                ) == MessageBoxResult.Yes
                            )
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                {
                                    FileName = path,
                                    UseShellExecute = true
                                });
                            }
                            App.Current.Shutdown(errorCode);
                            this.Close();
                        });
                        return;
                    }
                }
                Dispatcher.Invoke(() =>
                {
                    App.Current.Shutdown(0);
                    this.Close();
                });
                return;
            });
        }
        
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            cancellationTokenSource.Cancel();
            LogTextBox.AppendText("Conversion cancelled by user.\r\n");
            LogTextBox.ScrollToEnd();
            Console.WriteLine("Conversion cancelled by user.");
            App.Current.Shutdown(1);
            this.Close();
        }


    }
}
