using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Resources;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;


namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class Validation : Form
    {
        private ResourceManager resManager;
        String outputPath;

        /*Constructor*/
        public Validation(string label, string details, String outputFile, ResourceManager manager)
        {
            InitializeComponent();
            this.resManager = manager;
            outputPath = outputFile;
            lbl_Information.Text = manager.GetString(label);
            //string text = manager.GetString(details);
            details = details.Replace('\r', ' ');
            String[] str = details.Split('\n');
            for (int i = 0; i < str.Length; i++)
            {
                listBox_Validation.Items.Insert(i, str[i]);

            }
        }

        /*Function to show Help file*/
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm");
        }

        /*Function to close UI*/
        private void btn_OK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /*Function to save Log file*/
        private void saveloglinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog {
                Filter = "Text files (*.txt)|*.txt",
                FilterIndex = 1,
                CheckPathExists = true,
                DefaultExt = ".txt",
                FileName = Path.GetDirectoryName(outputPath) + "\\LogFile.txt"
            };
            sfd.ShowDialog();
            StreamWriter writer = new StreamWriter(sfd.FileName);
            writer.Write("Validation errors occured during Translation to " + outputPath.Substring(outputPath.LastIndexOf("\\") + 1) + ":");
            writer.Write(Environment.NewLine);
            writer.Write("**************************************");
            writer.Write(Environment.NewLine);
            for (int i = 0; i < listBox_Validation.Items.Count; i++)
            {
                writer.Write(listBox_Validation.Items[i].ToString());
                writer.Write(" ");
                writer.Write(Environment.NewLine);
            }
            writer.Close();
        }
     }
}