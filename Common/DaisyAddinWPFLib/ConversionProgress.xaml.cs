using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2;
using System;
using System.Threading;
using System.Windows;
namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Progress dialog singleton for conversion operations.
    /// </summary>
    public partial class ConversionProgress : Window
    {
        private string CurrentProgressMessage = "";
        private int StepIncrement = 1;
        private ConversionProgress()
        {
            InitializeComponent();
        }

        // for thread safety
        private static readonly object padlock = new object();
        private static ConversionProgress instance = null;
        private static Thread dialogThread = null;
        public static ConversionProgress Instance {
            get
            {
                lock (padlock) {
                    if (instance == null || (instance != null && !instance.IsVisible)) {
                        
                        try {
                            instance = new ConversionProgress();
                            instance.Show();
                            //dialogThread = new Thread(() => {
                            //    instance = new ConversionProgress();
                            //    instance.ShowDialog();
                            //});
                            //dialogThread.SetApartmentState(ApartmentState.STA);
                            //dialogThread.IsBackground = true;
                            //dialogThread.Start();
                            ////dialogThread.Join();
                            //while (instance == null || !instance.IsLoaded) {
                            //    Thread.Sleep(100); // Wait for the dialog to be initialized
                            //}
                        }
                        catch (Exception ex) {
                            throw new Exception("An error occured while initializing progress dialog", ex);
                        }
                    }
                    return instance;
                }
            }
        }




        // For external thread calls
        public delegate void DelegatedAddMessage(string message, bool isProgress = true);

        public void AddMessage(string message, bool isProgress = true)
        {
            if (this.Dispatcher.CheckAccess() == false) {
                this.Dispatcher.Invoke(new DelegatedAddMessage(AddMessage), message, isProgress);
                return;
            }
            string[] lines = message.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines) {
                if (isProgress) {
                    CurrentProgressMessage = line;
                    CurrentAction.Content = CurrentProgressMessage;
                    ProgressionTracker.Value += StepIncrement;
                } else {
                    CurrentAction.Content = CurrentProgressMessage + " - " + line;
                }
                MessageTextArea.AppendText(line + "\r\n");
            }
            //CurrentAction.UpdateLayout();
            //ProgressionTracker.UpdateLayout();
            //MessageTextArea.UpdateLayout();
            MessageTextArea.ScrollToEnd();
            this.UpdateLayout();
        }

        private delegate void DelegatedInitializeProgress(string message = "", int maximum = 1, int step = 1);

        /// <summary>
        /// Prepare the progress bar
        /// </summary>
        /// <param name="message">Progress initialization message</param>
        /// <param name="maximum">maximum value of the progression (number of step expected)</param>
        /// <param name="step">step increment (default to one)</param>
        public void InitializeProgress(string message = "", int maximum = 1, int step = 1)
        {
            if (this.Dispatcher.CheckAccess() == false) {
                this.Dispatcher.Invoke(new DelegatedInitializeProgress(InitializeProgress), message, maximum, step);
                return;
            }
            CurrentProgressMessage = message;
            CurrentAction.Content = CurrentProgressMessage;
            this.MessageTextArea.AppendText((message.EndsWith("\n") ? message : message + "\r\n"));
            ProgressionTracker.Maximum = maximum;
            StepIncrement = step;
            ProgressionTracker.Value = 0;
            //CurrentAction.UpdateLayout();
            //ProgressionTracker.UpdateLayout();
            //MessageTextArea.UpdateLayout();
            this.UpdateLayout();
        }

        private delegate int DelegatedProgress();

        /// <summary>
        /// Increment the progress bar
        /// </summary>
        /// <returns></returns>
        public int Progress()
        {
            if (this.Dispatcher.CheckAccess() == false) {
                return (int)this.Dispatcher.Invoke(new DelegatedProgress(Progress));
            }
            ProgressionTracker.Value += StepIncrement;
            return (int)ProgressionTracker.Value;

        }

        private void MessageTextArea_TextChanged(object sender, EventArgs e)
        {
            MessageTextArea.SelectionStart = MessageTextArea.Text.Length;
            MessageTextArea.ScrollToEnd();
        }


        public delegate void CancelClickListener();

        private event CancelClickListener cancelButtonClicked;

        public void setCancelClickListener(CancelClickListener cancelAction)
        {
            cancelButtonClicked = cancelAction;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (cancelButtonClicked != null) cancelButtonClicked();
            this.Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (cancelButtonClicked != null) cancelButtonClicked();
        }

    }
}
