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

        public ConversionProgress() {
            InitializeComponent();
        }

        public void addMessage(string message) {
            this.MessageTextArea.Text = this.MessageTextArea.Text + ( message.EndsWith("\n") ? message : message + "\r\n");
            this.MessageTextArea.Refresh();
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


    }
}
