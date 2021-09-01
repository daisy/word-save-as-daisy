using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Resources;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class Fidility : Form
    {
        private ResourceManager resManager;
        String outputPath, strResource;

        /*Constructor to set the DataGrid View properties*/
        public Fidility(string label, ArrayList details, String outputFile, ResourceManager manager)
        {
            InitializeComponent();
            this.resManager = manager;
            lbl_Information.Text = manager.GetString(label);
            outputPath = outputFile;
            StringBuilder bld = new StringBuilder();

            DataGridViewLinkColumn llc = new DataGridViewLinkColumn();
            DataGridViewTextBoxColumn l1 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn l2 = new DataGridViewTextBoxColumn();
            l1.Width = 22;
            l2.Width = 330;
            llc.Width = 70;

            llc.LinkBehavior = LinkBehavior.NeverUnderline;
            llc.Resizable = DataGridViewTriState.False;
            l1.Resizable = DataGridViewTriState.False;
            l2.Resizable = DataGridViewTriState.False;

            this.dataGrid_Fidility.Columns.Add(l1);
            this.dataGrid_Fidility.Columns.Add(l2);
            this.dataGrid_Fidility.Columns.Add(llc);

            for (int i = 0; i < details.Count; i++)
            {
                this.dataGrid_Fidility.Rows.Add(i + 1 + ".", details[i].ToString(), "More Info");
                this.dataGrid_Fidility.Rows[i].Resizable = DataGridViewTriState.False;
            }
            this.dataGrid_Fidility.ReadOnly = true;
            this.dataGrid_Fidility.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        /// <summary>
        /// Function to Show the Help File
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (this.dataGrid_Fidility.Columns[e.ColumnIndex] is DataGridViewLinkColumn)
            {
                String[] split;
                String str = this.dataGrid_Fidility.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                if (str.Contains("for"))
                {
                    str = str.Replace("for", "|for");
                    split = str.Split('|');
                    str = split[0].TrimEnd();
                }
                strResource = resManager.GetString(str);
                if (strResource != null)
                    Help.ShowHelp(dataGrid_Fidility, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.KeywordIndex, strResource);
                else
                    Help.ShowHelp(dataGrid_Fidility, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.TableOfContents);
            }
        }

        /// <summary>
        /// Function to Close the UI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_OK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Function to show Help file on Click of Enter button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid_Fidility_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.Equals('\r'))
            {
                DataGridViewCell cell = dataGrid_Fidility.CurrentCell;
                if (this.dataGrid_Fidility.Columns[cell.ColumnIndex] is DataGridViewLinkColumn)
                {
                    String str = this.dataGrid_Fidility.Rows[cell.RowIndex].Cells[cell.ColumnIndex - 1].Value.ToString();
                    strResource = resManager.GetString(str);
                    if (strResource != null)
                        Help.ShowHelp(dataGrid_Fidility, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.KeywordIndex, strResource);
                    else
                        Help.ShowHelp(dataGrid_Fidility, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.TableOfContents);
                }
            }
        }

        /// <summary>
        /// Function to save Log file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lLbl_Log_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text files (*.txt)|*.txt";
            sfd.FilterIndex = 1;
            sfd.CheckPathExists = true;
            sfd.DefaultExt = ".txt";
            sfd.FileName = Path.GetDirectoryName(outputPath) + "\\LogFile.txt";
            sfd.ShowDialog();
            StreamWriter writer = new StreamWriter(sfd.FileName);
            writer.Write("Fidelity Loss during Translation " + outputPath.Substring(outputPath.LastIndexOf("\\") + 1) + ":");
            writer.Write(Environment.NewLine);
            writer.Write("**************************************");
            writer.Write(Environment.NewLine);
            for (int i = 0; i < dataGrid_Fidility.Rows.Count; i++)
            {
                for (int j = 0; j < dataGrid_Fidility.Columns.Count - 1; j++)
                {
                    writer.Write(this.dataGrid_Fidility.Rows[i].Cells[j].Value);
                    writer.Write(" ");
                }
                writer.Write(Environment.NewLine);
            }
            writer.Close();
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}