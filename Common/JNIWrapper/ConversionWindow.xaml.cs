using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

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
            Task.Run(() => JNITaskRunner.RunConversion(scriptName, options, cancellationTokenSource.Token,
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
                        ConversionProgressPercentage.Text = $"{(double)(p * 100) / (double)t} %";
                    });
                }),
                properties: properties
            ).ContinueWith(t =>
            {
                if (t.IsCanceled)
                {
                    Dispatcher.Invoke(() =>
                    {
                        ConversionProgressText.Text = "Conversion was cancelled.";
                        LogTextBox.AppendText("Conversion was cancelled." + "\r\n");
                        LogTextBox.ScrollToEnd();
                        Application.Current.Shutdown(1);
                    });
                }
                else if (t.IsFaulted)
                {
                    Dispatcher.Invoke(() =>
                    {
                        string path = Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-unknowerror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                        AggregateException aggex = t.Exception;
                        foreach (Exception iex in aggex.InnerExceptions)
                        {
                            StringBuilder errorLog = new StringBuilder($"An error occured running {scriptName} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}\r\n");
                            LogTextBox.AppendText($"An error occured running {scriptName} with options : {string.Join(" ", options.Select(kv => "--" + kv.Key + " \"" + kv.Value + "\""))}" + "\r\n");
                            StringBuilder errorMessage = new StringBuilder();
                            LogTextBox.ScrollToEnd();
                            var ex = iex;
                            int errorCode = 1;
                            while (ex != null)
                            {
                                errorMessage.AppendLine(ex.Message);
                                LogTextBox.AppendText($"{ex.Message}\r\n");
                                LogTextBox.ScrollToEnd();
                                errorLog.AppendLine(ex.Message);
                                errorLog.AppendLine(ex.StackTrace);
                                ex = ex.InnerException;
                            }
                            if (iex is OperationCanceledException)
                            {
                                LogTextBox.AppendText("Conversion was cancelled." + "\r\n");
                                LogTextBox.ScrollToEnd();
                                Console.WriteLine("Conversion was cancelled.");

                            }
                            else if (iex is JobException jex)
                            {
                                path = Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-joberror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                                errorCode = 2;
                            }
                            else
                            {
                                path = Path.Combine(JNITaskRunner.JNILogsPath, $"{scriptName}-criticalerror-{DateTime.Now.ToString("yyyyMMddHHmmss")}.log");
                                errorCode = 3;
                            }
                            if (errorCode == 1)
                            {
                                // Cancel action, don't report anything, just shutdown
                                Application.Current.Shutdown(errorCode);
                            }
                            else
                            {
                                File.WriteAllText(path, errorLog.ToString());

                                if (MessageBox.Show(
                                        $"The following error occured during the conversion:" +
                                        $"\r\n{errorMessage}" +
                                        $"\r\n Do you want to open the error log at {path} for more details?",
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
                                Application.Current.Shutdown(errorCode);
                            }
                        }
                    });
                }
                else
                {
                    Dispatcher.Invoke(() =>
                    {
                        ConversionProgressText.Text = "Conversion completed successfully.";
                        LogTextBox.AppendText("Conversion completed successfully." + "\r\n");
                        LogTextBox.ScrollToEnd();
                        Application.Current.Shutdown(0);
                    });
                }

            }));

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


    }
}
