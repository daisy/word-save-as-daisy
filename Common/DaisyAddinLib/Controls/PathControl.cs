using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class PathControl : BaseUserControl
    {

        private string selectedPath;
        private PathDataType linkedPathData;

        public PathControl()
        {
            InitializeComponent();
        }

        public PathControl(ScriptParameter p) : base(p)
        {
            InitializeComponent();
            setLinkedParameter(p);

        }

        public override void setLinkedParameter(ScriptParameter p)
        {
            base.setLinkedParameter(p);
            linkedPathData = (PathDataType)p.ParameterDataType;
            SelectedPath = p.ParameterValue.ToString();

            parameterValue.AccessibleName = p.NiceName;
            parameterValue.AccessibleDescription = p.Description;
            descriptionTooltip.SetToolTip(parameterValue, p.Description);

            browseButton.AccessibleName = "Select " + p.NiceName;
            browseButton.AccessibleDescription = p.Description;
        }


        public string SelectedPath
        {
            get
            {
                return selectedPath;
            }
            set
            {
                if (value != null)
                {
                    selectedPath = value;
                    this.parameterValue.SelectionStart = 0;
                    parameterValue.Text = value;
                }
            }
        }

        public new string ParameterName { get => this.parameterNiceName.Text; set => this.parameterNiceName.Text = value; }

        private void OnClickBrowseButton(object sender, EventArgs e)
        {
            if (linkedPathData.IsFile)
            {
                if (linkedPathData.IsInput) UpdatePathTextboxFromOpenFileDialog();
                else UpdatePathTextboxFromSaveFileDialog();
            }
            else
            {
                UpdatePathTextboxFromFolderBrowserDialog();
            }
        }

        private void UpdatePathTextboxFromOpenFileDialog()
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                parameterValue.Text = openFileDialog1.FileName;
            }
        }

        private void UpdatePathTextboxFromSaveFileDialog()
        {
            if (!string.IsNullOrEmpty(linkedPathData.FileExtension))
                saveFileDialog1.Filter = linkedPathData.FileExtension + "|*" + linkedPathData.FileExtension;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                parameterValue.Text = saveFileDialog1.FileName;
            }
        }

        public override void UpdateScriptParameterValue()
        {
            if ((parameterValue.Text != null && parameterValue.Text.TrimEnd() != ""))
            {
                selectedPath = parameterValue.Text;
                if (linkedPathData.IsOutput)
                {
                    try
                    {
                        if (linkedPathData.IsDirectory) CheckForOutputDirectory();
                        else File.CreateText(parameterValue.Text).Close();
                    }
                    catch (System.Exception ex)
                    {
                        AddinLogger.Error(ex);
                        return;
                    }
                }
                // update    parameter
                try
                {
                    linkedPathData.Value = parameterValue.Text;
                }
                catch (System.Exception ex)
                {
                    AddinLogger.Error(ex);
                }
            }
            return;// null value check ends
        }


        private bool CheckForOutputDirectory()
        {
            if (!Directory.Exists(parameterValue.Text))
            {
                Directory.CreateDirectory(parameterValue.Text);
                return false;
            }
            else
                return false;
        }


        private string GetPhysicalPath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (parameterValue.Text != "")
                fbd.SelectedPath = parameterValue.Text;
            fbd.ShowDialog();
            return fbd.SelectedPath;
            // return the path in which the user wants to create the file
        }
        private void UpdatePathTextboxFromFolderBrowserDialog()
        {
            String path = GetPhysicalPath();
            folderBrowserDialog1.ShowNewFolderButton = true;
            if (path != "")
            {
                parameterValue.Text = path;
            }

        }

        //private void PathControl_Load(object sender, EventArgs e)
        //{
        //    if(this.Parent != null)
        //    {
        //        this.Width = this.Parent.Width;
        //        InputPanel.Width = this.Width - 20;
        //    }
        //}
    }
}
