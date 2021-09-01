using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms.Controls {
	public partial class PathBrowserControl : BaseUserControl
	{
		private string m_SelectedPath;
		private string m_TextLabel;
		private ScriptParameter m_Parameter;
		private PathDataType m_PathData;

		public PathBrowserControl() { InitializeComponent(); }

		public PathBrowserControl(ScriptParameter p, string SelectedPath,  string strBrtextBox)
			: this()
		{
			m_SelectedPath = SelectedPath;
			m_Parameter = p;
			//m_PreDestination = strBrtextBox;

			m_PathData = (PathDataType)p.ParameterDataType;

			int wdiff = mNiceNameLabel.Width;
			mNiceNameLabel.Text = p.NiceName;

			int x = mNiceNameLabel.Width;
			wdiff -= mNiceNameLabel.Width;
			if (wdiff < 0)
			{
				Point location = mNiceNameLabel.Location;
				Width -= wdiff;
				mNiceNameLabel.Location = location;

			}
			else
			{
				mNiceNameLabel.Location = new Point(mNiceNameLabel.Location.X - wdiff, mNiceNameLabel.Location.Y);
			}
			mTextBox.AccessibleName = p.Description;
			mTextBox.Text = strBrtextBox;
			mTextBox.SelectionStart = 0;

			try
			{
				if (m_PathData.isInputOrOutput == PathDataType.InputOrOutput.input)
				{
					mTextBox.Text = SelectedPath;
				}
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}
		public PathBrowserControl(ScriptParameter p, string SelectedPath)
			: this()
		{
			m_SelectedPath = SelectedPath;
			m_Parameter = p;
			//m_PreDestination = strBrtextBox;

			m_PathData = (PathDataType)p.ParameterDataType;

			int wdiff = mNiceNameLabel.Width;
			mNiceNameLabel.Text = p.NiceName;

			int x = mNiceNameLabel.Width;
			wdiff -= mNiceNameLabel.Width;
			if (wdiff < 0)
			{
				Point location = mNiceNameLabel.Location;
				Width -= wdiff;
				mNiceNameLabel.Location = location;

			}
			else
			{
				mNiceNameLabel.Location = new Point(mNiceNameLabel.Location.X - wdiff, mNiceNameLabel.Location.Y);
			}
			mTextBox.AccessibleName = p.Description;
			try
			{
				if (m_PathData.isInputOrOutput == PathDataType.InputOrOutput.input)
				{
					mTextBox.Text = SelectedPath;
				}
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}
		}

		public override string Text
		{
			get { return m_TextLabel; }
			set
			{
				if (value != null)
				{
					mNiceNameLabel.Text = value;
					mTextBox.AccessibleName = value;
				}
			}
		}



		public string SelectedPath
		{
			get
			{
				return m_SelectedPath;
			}
			set
			{
				if (value != null)
				{
					mTextBox.Text = value;

				}
			}
		}

		private void btnBrowse_Click(object sender, EventArgs e)
		{
			if (m_PathData.IsFileOrDirectory == PathDataType.FileOrDirectory.File)
			{
				if (m_PathData.isInputOrOutput == PathDataType.InputOrOutput.input)
					UpdatePathTextboxFromOpenFileDialog();
				else if (m_PathData.isInputOrOutput == PathDataType.InputOrOutput.output)
					UpdatePathTextboxFromSaveFileDialog();
			}
			else if (m_PathData.IsFileOrDirectory == PathDataType.FileOrDirectory.Directory)
			{
				UpdatePathTextboxFromFolderBrowserDialog();
			}
		}

		private void UpdatePathTextboxFromOpenFileDialog()
		{
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				mTextBox.Text = openFileDialog1.FileName;
			}
		}

		private void UpdatePathTextboxFromSaveFileDialog()
		{
			if (!string.IsNullOrEmpty(m_PathData.FileExtenssion))
				saveFileDialog1.Filter = m_PathData.FileExtenssion + "|*" + m_PathData.FileExtenssion;

			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				mTextBox.Text = saveFileDialog1.FileName;
			}
		}

		//private void UpdatePathTextboxFromFolderBrowserDialog()
		//{
		//    folderBrowserDialog1.ShowNewFolderButton = true;
		//    if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
		//    {
		//        mTextBox.Text = folderBrowserDialog1.SelectedPath;
		//    }

		//}

		public override void UpdateScriptParameterValue()
		{
			if ((mTextBox.Text != null && mTextBox.Text.TrimEnd() != ""))
			{
				m_SelectedPath = mTextBox.Text;
				if (m_PathData.isInputOrOutput == PathDataType.InputOrOutput.output)
				{
					try
					{
						if (m_PathData.IsFileOrDirectory == PathDataType.FileOrDirectory.Directory)
							CheckForOutputDirectory();
						else if (m_PathData.IsFileOrDirectory == PathDataType.FileOrDirectory.File)
						{
							File.CreateText(mTextBox.Text).Close();
						}
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
					m_PathData.Value = mTextBox.Text;
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
			if (!Directory.Exists(mTextBox.Text))
			{
				Directory.CreateDirectory(mTextBox.Text);
				return false;
			}
			else
				return false;
		}

		private void mNiceNameLabel_Click(object sender, EventArgs e)
		{

		}

		private void PathBrowserControl_Load(object sender, EventArgs e)
		{

		}
		private string GetPhysicalPath()
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			if (mTextBox.Text != "")
				fbd.SelectedPath = mTextBox.Text;
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
				mTextBox.Text = path;
			}

		}

		// end of dir check function
	}
}
