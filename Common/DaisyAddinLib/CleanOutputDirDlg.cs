using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.Conversion
{
	public partial class CleanOutputDirDlg : Form
	{
		private string[] _dirictoriesToClean = new string[] { "z3986", "daisy202" };
		private readonly string _outputDir;
		private readonly string _outputFile;

		public CleanOutputDirDlg()
			: this(string.Empty, string.Empty)
		{
		}

		public CleanOutputDirDlg(string outputDir, string outputFile)
		{
			InitializeComponent();
			_outputDir = outputDir;
			_outputFile = outputFile;
			OutputDir = outputDir;
		}

		public string OutputDir { get; private set; }

		public DialogResult Clean(Form parent)
		{
			if (!Directory.Exists(_outputDir))
				return DialogResult.OK;

			DirectoryInfo directoryInfo = new DirectoryInfo(_outputDir);

			DirectoryInfo[] subDirectories = directoryInfo.GetDirectories();

			bool isNeedInCleaning = false;

			foreach (DirectoryInfo subDirectory in subDirectories)
			{
				if (IsDirectoryShouldBeDeleted(subDirectory))
				{
					isNeedInCleaning = true;
					break;
				}
			}

			if (!isNeedInCleaning)
				return DialogResult.OK;

			return this.ShowDialog(parent);
		}



		private void cleanButton_Click(object sender, EventArgs e)
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(_outputDir);
			DirectoryInfo[] subDirectories = directoryInfo.GetDirectories();
			try
			{
				foreach (DirectoryInfo subDirectory in subDirectories)
				{
					if (IsDirectoryShouldBeDeleted(subDirectory))
						Directory.Delete(subDirectory.FullName, true);
				}

				DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				DialogResult = DialogResult.Cancel;
			}
		}

		private bool IsDirectoryShouldBeDeleted(DirectoryInfo directoryInfo)
		{
			foreach (string directoryName in _dirictoriesToClean)
			{
				if (directoryName.Equals(directoryInfo.Name, StringComparison.CurrentCultureIgnoreCase))
					return true;
			}
			return false;
		}

		private void continueButton_Click(object sender, EventArgs e)
		{
			try
			{
				OutputDir = Path.Combine(OutputDir, string.Format("output_{0}", DateTime.Now.ToString("dd_MM_yyyy_HHmmss")));
				Directory.CreateDirectory(OutputDir);
				DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
				DialogResult = DialogResult.Cancel;
			}
		}
	}
}
