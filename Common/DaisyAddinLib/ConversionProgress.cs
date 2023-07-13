using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY {
    public partial class ConversionProgress : Form {

        private string CurrentProgressMessage = "";
        public ConversionProgress() {
            InitializeComponent();
        }

        // For external thread calls
        public delegate void DelegatedAddMessage(string message, bool isProgress = true);

        public void AddMessage(string message, bool isProgress = true) {
            if (this.InvokeRequired) {
                this.BeginInvoke(new DelegatedAddMessage(AddMessage), message, isProgress);
            } else {
                string[] lines = message.Split(new char[] {'\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    if (isProgress)
                    {
                        CurrentProgressMessage = line;
                        LastMessage.Text = CurrentProgressMessage;
                        ConversionProgressBar.PerformStep();
                    }
                    else
                    {
                        LastMessage.Text = CurrentProgressMessage + " - " + line;
                    }
                    this.MessageTextArea.BeginInvoke((MethodInvoker)delegate { this.MessageTextArea.AppendText(line + "\r\n"); });
                    //this.MessageTextArea.AppendText(line + "\r\n");
                    //this.MessageTextArea.Text += line + "\r\n";
                }
                
            }
            
        }

        private delegate void DelegatedInitializeProgress(string message = "", int maximum = 1, int step = 1);

        /// <summary>
        /// Prepare the progress bar
        /// </summary>
        /// <param name="message">Progress initialization message</param>
        /// <param name="maximum">maximum value of the progression (number of step expected)</param>
        /// <param name="step">step increment (default to one)</param>
        public void InitializeProgress(string message = "", int maximum = 1, int step = 1) {
            if (this.InvokeRequired) {
                this.Invoke(new DelegatedInitializeProgress(InitializeProgress), message, maximum, step);
            } else {
                if(!this.IsDisposed)
                {
                    CurrentProgressMessage = message;
                    LastMessage.Text = CurrentProgressMessage;
                    this.MessageTextArea.BeginInvoke(
                        (MethodInvoker)delegate {
                            this.MessageTextArea.AppendText((message.EndsWith("\n") ? message : message + "\r\n"));
                        }
                    );
                    ConversionProgressBar.Maximum = maximum;
                    ConversionProgressBar.Step = step;
                    ConversionProgressBar.Value = 0;
                }
            }
            

        }

        private delegate int DelegatedProgress();

        /// <summary>
        /// Increment the progress bar
        /// </summary>
        /// <returns></returns>
        public int Progress() {
            if (this.InvokeRequired) { return (int)this.Invoke(new DelegatedProgress(Progress)) ; } else
            {
                ConversionProgressBar.PerformStep();
                return ConversionProgressBar.Value;
            }
            
        }

        private void MessageTextArea_TextChanged(object sender, EventArgs e) {
            MessageTextArea.SelectionStart = MessageTextArea.Text.Length;
            MessageTextArea.ScrollToCaret();
        }


        public delegate void CancelClickListener();

        private event CancelClickListener cancelButtonClicked;

        public void setCancelClickListener(CancelClickListener cancelAction) {
            cancelButtonClicked = cancelAction;
        }

        private void CancelButton_Click(object sender, EventArgs e) {
            if (cancelButtonClicked != null) cancelButtonClicked();
            this.Close();
        }

        private void ConversionProgress_FormClosing(Object sender, FormClosingEventArgs e) {
            if (cancelButtonClicked != null) cancelButtonClicked();
        }

        private void ShowHideDetails_Click(object sender, EventArgs e) {
            if(MessageTextArea.Visible == false) {
                MessageTextArea.Visible = true;
                this.Height = 300;
                MessageTextArea.Height = 135;
                MessageTextArea.SelectionStart = MessageTextArea.Text.Length;
                MessageTextArea.ScrollToCaret();
                ShowHideDetails.Text = "Hide details << ";
            } else {
                MessageTextArea.Visible = false;
                this.Height = 300 - 135;
                MessageTextArea.Height = 0;
                ShowHideDetails.Text = "Show details >> ";
            }
        }

    }
}
