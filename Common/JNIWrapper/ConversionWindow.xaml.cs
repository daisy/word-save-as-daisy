using System;
using System.Collections.Generic;
using System.Threading;
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

            JNITaskRunner.LaunchConversion(scriptName, options, cancellationTokenSource.Token,
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
                progress: ((p,t) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            ProgressBar.Maximum = t;
                            ProgressBar.Value = p;
                        });
                    }),
                properties: properties
            ).ContinueWith(t =>
            {
                Dispatcher.Invoke(() =>
                {
                    Application.Current.Shutdown(t.Result);
                });
               
            });

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
