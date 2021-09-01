using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class MasterSubValidation : Form
    {
        public MasterSubValidation(String details, String errorType)
        {
            InitializeComponent();

            if (errorType == "Validation")
                this.Text = "SaveAsDAISY - Validation Error";
            else if (errorType == "Success")
                this.Text = "SaveAsDAISY - Success";
            else
                this.Text = "SaveAsDAISY";
            this.richTextBox1.Text = details;
        }

        /*Function to close UI*/
        private void OK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /*Function to close UI*/
        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /*Function to save Log File*/
        private void saveloglinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files (*.txt)|*.txt";
            sfd.FilterIndex = 1;
            sfd.CheckPathExists = true;
            sfd.DefaultExt = ".txt";
            sfd.ShowDialog();
            StreamWriter writer = new StreamWriter(sfd.FileName);
            writer.Write("Log File:");
            writer.Write(Environment.NewLine);
            writer.Write("**********");
            writer.Write(Environment.NewLine);
            writer.Write(richTextBox1.Text.Replace("\n", Environment.NewLine));
            writer.Close();
        }

        /*Function to show Help file*/
        private void linkLabel_Guidelines_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm");
        }
    }
}