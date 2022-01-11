using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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

        public void addMessage(string message, bool isProgress = true) {
            if (isProgress) {
                CurrentProgressMessage = message;
                LastMessage.Text = CurrentProgressMessage;
                ConversionProgressBar.PerformStep();
            } else {
                LastMessage.Text = CurrentProgressMessage + " - " + message;
            }
            this.MessageTextArea.Text += ( message.EndsWith("\n") ? message : message + "\r\n");
        }


        /// <summary>
        /// Prepare the progress bar
        /// </summary>
        /// <param name="message">Progress initialization message</param>
        /// <param name="maximum">maximum value of the progression (number of step expected)</param>
        /// <param name="step">step increment (default to one)</param>
        public void InitializeProgress(string message = "", int maximum = 1, int step = 1) {
            CurrentProgressMessage = message;
            LastMessage.Text = CurrentProgressMessage;
            this.MessageTextArea.Text += (message.EndsWith("\n") ? message : message + "\r\n");
            ConversionProgressBar.Maximum = maximum;
            ConversionProgressBar.Step = step;
            ConversionProgressBar.Value = 0;

        }

        public int Progress() {
            ConversionProgressBar.PerformStep();
            return ConversionProgressBar.Value;
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
