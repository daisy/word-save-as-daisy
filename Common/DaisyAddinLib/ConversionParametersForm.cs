/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Drawing;
using System.IO;
using System.Xml;
using System.Resources;
using System.Reflection;
using System.Collections;
using System.IO.Packaging;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion;
using System.Diagnostics;
using Daisy.SaveAsDAISY.Forms;
using Daisy.SaveAsDAISY.Forms.Controls;

namespace Daisy.SaveAsDAISY.Conversion
{
	/**
	 * Form to graphicaly complete a conversion parameters set
	 */
	public partial class ConversionParametersForm : Form
	{
		int translateFlag = 0;
		XmlDocument XmlPackage;
		private ResourceManager labels;
		Hashtable parametersHash = new Hashtable();
		string officeVersion = "";
		PackageRelationship packRelationship = null;
		string inputFileName, embedFilePath, tempInput, trackChangeFlag, outputFilePath, masterSubFlag;
		const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		String uId = "";
		string btnID = "";
		private string mInputPath;
		private bool useAScript = false;
		String strBrtextBox = "";
		TableLayoutPanel oTableLayoutPannel = new TableLayoutPanel();

		public ConversionParameters UpdatedConversionParameters { get; private set; }


		/// <summary>
		/// Returns Hash Table having information about Title,Creator,Publisher,UID
		/// </summary>
		public Hashtable ParametersHash { get { return parametersHash; } }

		/// <summary>
		/// Return Title information 
		/// </summary>
		public string GetTitle { get { return tBx_Title.Text; } }

		public ScriptParser getParser { get { return UpdatedConversionParameters.PostProcessSettings; } }

		/// <summary>
		/// Return Creator information 
		/// </summary>
		public string GetCreator { get { return tBx_Creator.Text; } }

		/// <summary>
		/// Returns the value whether to translate the current document
		/// along with sub documents
		/// </summary>
		public string MastersubValue { get { return masterSubFlag; } }


		/// <summary>
		/// Returns the path of the output file
		/// </summary>
		public string OutputFilepath { get { return outputFilePath; } }

		public string PipeOutput { get { return strBrtextBox; } }


		/// <summary>
		/// Default form for converting a word file to DTbook XML
		/// </summary>
		/// <param name="conversion"></param>
		/// <param name="labelsManager"></param>
		public ConversionParametersForm(DocumentParameters document, ConversionParameters conversion, ResourceManager labelsManager = null) {
			// Copy current conversion settings
			UpdatedConversionParameters = conversion.usingMainDocument(document);

			btnID = conversion.ControlName;
			tempInput = document.CopyPath ?? document.InputPath;
			this.labels = labelsManager ?? new ResourceManager(
					"DaisyAddinLib.resources.Labels",
					Assembly.GetExecutingAssembly()
				);
			//this.officeVersion = conversion.Version;
			//this.masterSubFlag = conversion.ParseSubDocuments;

			// if a script is defined in the parameters
			useAScript = UpdatedConversionParameters.ScriptPath != null && UpdatedConversionParameters.ScriptPath.Length > 0;
			if (useAScript) {
				mInputPath = document.InputPath;
				if (UpdatedConversionParameters.PostProcessSettings == null) 
					UpdatedConversionParameters.PostProcessSettings = new ScriptParser(UpdatedConversionParameters.ScriptPath);

			} else {
				inputFileName = document.InputPath;
			}
			
			InitializeComponent();
		}

		/// <summary>
		/// Function to select Output folder 
		/// </summary>
		/// <returns>Output filename chosen by the user</returns>
		private string GetPhysicalPath()
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			fbd.Description = labels.GetString("Destination");
			if (tBx_Browse.Text != "")
				fbd.SelectedPath = tBx_Browse.Text;
			else
				fbd.SelectedPath = Path.GetDirectoryName(inputFileName);

			fbd.ShowDialog();
			return fbd.SelectedPath;
			// return the path in which the user wants to create the file
		}

		/// <summary>
		/// Function to Get the data of a particular file
		/// </summary>
		/// <param name="fileName">File name </param>
		/// <returns></returns>
		public Stream GetInnerStream(String fileName)
		{
			string location = ".Resources.";
			embedFilePath = this.GetType().Namespace + location;
			String final = String.Concat(embedFilePath, fileName);
			return Assembly.GetExecutingAssembly().GetManifestResourceStream(final);
		}

		/*private bool IsTranslateToSingleDaisy
		{
			get { return AddInHelper.buttonIsSingleWordToXMLConversion(btnID) || !ConverterHelper.PipelineIsInstalled(); }
		}*/

		public void Translate()
		{
			bool isValidInputs = !useAScript
							? ValidateForSingleDaisyTranslate()
							: ValidateForFullDaisyTranslate();

			if (isValidInputs)
			{
				if (!useAScript) outputFilePath = tBx_Browse.Text;

				UpdatePopulateOutputXml(!useAScript ? tBx_Browse.Text : strBrtextBox);

				parametersHash = updateParametersHash();

				translateFlag = 1;
				this.Close();
			}

		}

		private bool ValidateForFullDaisyTranslate()
		{
			string fileName = Path.GetFileNameWithoutExtension(mInputPath);
			string otpfileName = fileName + ".xml";
			tBx_Browse.Text = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY\";
			if (tBx_Title.Text.TrimEnd() == "")
			{
				MessageBox.Show(labels.GetString("Title"), labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				tBx_Title.Focus();
				return false;
			}

			for (int i = 0; i < mLayoutPanel.Controls.Count; i++)
			{
				if (mLayoutPanel.Controls[i] is BaseUserControl)
				{

					((BaseUserControl)mLayoutPanel.Controls[i]).UpdateScriptParameterValue();

				}
			}

			for (int i = 0; i < oTableLayoutPannel.Controls.Count; i++)
			{
                if (oTableLayoutPannel.Controls[i] is BaseUserControl)
				{
                    ((BaseUserControl)oTableLayoutPannel.Controls[i]).UpdateScriptParameterValue();
				}
			}

			foreach (ScriptParameter p in UpdatedConversionParameters.PostProcessSettings.ParameterList)
			{
				if (p.IsParameterRequired && (p.Name == "outputPath" || p.Name == "output"))
				{
					PathDataType pathDataType = p.ParameterDataType as PathDataType;
					if (pathDataType == null) continue;

					if (pathDataType.IsFileOrDirectory == PathDataType.FileOrDirectory.File)
					{
						try
						{
							FileInfo outputFileInfo = new FileInfo(p.ParameterValue);
							if (!string.IsNullOrEmpty(pathDataType.FileExtenssion) && !pathDataType.FileExtenssion.Equals(outputFileInfo.Extension, StringComparison.InvariantCultureIgnoreCase))
							{
								MessageBox.Show(string.Format("Please select {0} output file", pathDataType.FileExtenssion), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
								mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
								return false;
							}
							strBrtextBox = outputFileInfo.DirectoryName;
						}
						catch (ArgumentException ex)
						{
							AddinLogger.Error(ex);
							MessageBox.Show("Please select output file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
							return false;
						}

					}
					else
					{
						strBrtextBox = p.ParameterValue;
					}

					//TODO:::
					if (string.IsNullOrEmpty(strBrtextBox.TrimEnd()))
					{
						MessageBox.Show("Please select the Destination folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
						return false;
					}

					CleanOutputDirDlg cleanOutputDirDlg = new CleanOutputDirDlg(strBrtextBox, otpfileName);
					if (cleanOutputDirDlg.Clean(this) == DialogResult.Cancel)
					{
						mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
						return false;
					}
					if (strBrtextBox != cleanOutputDirDlg.OutputDir)
					{
						strBrtextBox = cleanOutputDirDlg.OutputDir;
						p.ParameterValue = cleanOutputDirDlg.OutputDir;
					}

					break;

				}
			}

			

			//if (File.Exists(strBrtextBox + "\\" + otpfileName))
			//{
			//    DialogResult objResult;
			//    objResult = MessageBox.Show(otpfileName + " already exists in the destination folder. Do you want to overwrite?", "Confirm File Replace", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
			//    if (objResult != DialogResult.Yes)
			//    {
			//        mLayoutPanel.Controls[0].Controls[0].Focus();
			//        return false;
			//    }
			//}

			return true;
		}

		private bool ValidateForSingleDaisyTranslate()
		{
			string fileName = Path.GetFileNameWithoutExtension(inputFileName);
			string otpfileName = fileName + ".xml";

			if (string.IsNullOrEmpty(tBx_Browse.Text.TrimEnd()))
			{
				MessageBox.Show(labels.GetString("ChoseDestination"), labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				tBx_Browse.Focus();
				return false;
			}

			if (Directory.Exists(tBx_Browse.Text) == false)
			{
				MessageBox.Show(string.Concat(tBx_Browse.Text, " ", "does not exist"), labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			if (tBx_Title.Text.TrimEnd() == "")
			{
				MessageBox.Show(labels.GetString("Title"), labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				tBx_Title.Focus();
				return false;
			}

			if (File.Exists(tBx_Browse.Text + "\\" + otpfileName))
			{
				DialogResult objResult;
				objResult = MessageBox.Show(otpfileName + " already exists in the destination folder. Do you want to overwrite?", "Confirm File Replace", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
				if (objResult != DialogResult.Yes)
				{
					tBx_Browse.Focus();
					return false;
				}
			}

			return true;
		}

		private Hashtable updateParametersHash()
		{

			UpdatedConversionParameters = UpdatedConversionParameters.withParameter("OutputFile", !useAScript ? tBx_Browse.Text : strBrtextBox)
				.withParameter("Title", tBx_Title.Text)
				.withParameter("Creator", tBx_Creator.Text)
				.withParameter("Publisher", tBx_Publisher.Text)
				.withParameter("UID", uId == tBx_Uid.Text ? "AUTO-UID-" + tBx_Uid.Text : tBx_Uid.Text)
				.withParameter("Subject", string.Empty)
				//.withParameter("TrackChanges", TrackChange()) // Done in preprocessing
				//.withParameter("Version", officeVersion)
				//.withParameter("MasterSubFlag", masterSubFlag)
				.withParameter("PipelineOutput", PipeOutput);
				//.withParameter("OutputFile", outputFilePath);
				//.withParameter("PostProcessSetting", Conversion.PostProcessSettings);


			return UpdatedConversionParameters.ConversionParametersHash;
		}

		private void UpdatePopulateOutputXml(string outputPath)
		{
			PrepopulateDaisyOutput output = new PrepopulateDaisyOutput(outputPath);
			output.Save();
		}

		/// <summary>
		/// Returns false when Form is diplayed to the user
		/// </summary>
		/// <returns>integer saying whether to continue or not</returns>
		public int DoTranslate()
		{
			this.ShowDialog();

			return translateFlag;
		}

		/// <summary>
		/// Function to Generate an unique ID
		/// </summary>
		/// <returns>ID</returns>
		public long GenerateId()
		{
			byte[] buffer = Guid.NewGuid().ToByteArray();
			return BitConverter.ToInt64(buffer, 0);
		}

		/// <summary>
		/// Function to check Track changes is on or Off for the current DOcument
		/// 
		/// </summary>
		/// <returns></returns>
		public String TrackChange()
		{
			Package pack = Package.Open(tempInput, FileMode.Open, FileAccess.ReadWrite);

			foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
			{
				packRelationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
			PackagePart mainPartxml = pack.GetPart(partUri);

			XmlPackage = new XmlDocument();
			XmlPackage.Load(mainPartxml.GetStream());
			pack.Close();

			NameTable nt = new NameTable();
			XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
			nsManager.AddNamespace("w", docNamespace);

			XmlNodeList listDel = XmlPackage.SelectNodes("//w:del", nsManager);
			XmlNodeList listIns = XmlPackage.SelectNodes("//w:ins", nsManager);

			if (listDel.Count > 0 || listIns.Count > 0)
			{
				DialogResult dr = MessageBox.Show(labels.GetString("TrackConfirmation"), "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

				if (dr == DialogResult.Yes)
					trackChangeFlag = "Yes";
				else
					trackChangeFlag = "No";
			}
			else
				trackChangeFlag = "NoTrack";

			return trackChangeFlag;
		}

		/// <summary>
		///  Function to check whether Current doc is having Sub documents or not
		///  To be removed : this is resolved in preprocessing phase of the document
		/// </summary>
		/// <returns></returns>
		public String MasterSubDecision()
		{
			Package pack = Package.Open(tempInput, FileMode.Open, FileAccess.ReadWrite);

			foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
			{
				packRelationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
			PackagePart mainPartxml = pack.GetPart(partUri);
			int cnt = 0;
			foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
			{
				packRelationship = searchRelation;
				if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
				{
					if (packRelationship.TargetMode.ToString() == "External")
					{
						cnt++;
					}
				}
			}
			pack.Close();

			if (cnt > 0)
			{
				DialogResult dr = MessageBox.Show(labels.GetString("MasterSubConfirmation"), "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

				if (dr == DialogResult.Yes)
					masterSubFlag = "Yes";
				else
					masterSubFlag = "No";
			}
			else
				masterSubFlag = "NoMasterSub";

			return masterSubFlag;
		}

		#region form & controls event handlers

		/// <summary>
		/// Function which loads the form 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void ConversionParametersFrom_Load(object sender, EventArgs e)
		{
			tBx_Uid.Text = GenerateId().ToString();
			uId = tBx_Uid.Text;

			tBx_Creator.Text = UpdatedConversionParameters.Creator;
			tBx_Title.Text = UpdatedConversionParameters.Title;
			tBx_Publisher.Text = UpdatedConversionParameters.Publisher;

			if (!useAScript)
			{
				PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();

				tBx_Browse.Text = prepopulateDaisyOutput != null
									  ? prepopulateDaisyOutput.OutputPath
									  : Path.GetDirectoryName(inputFileName);
				tBx_Browse.SelectionStart = 0;

				mLayoutPanel.Visible = false;
				oLayoutPanel.Visible = false;
				btn_OpenDetail.Visible = false;

				btn_OK.Location = new Point(grpBox_Properties.Width - 220, grpBox_Properties.Location.Y + grpBox_Properties.Height + 3);
				btn_Reset.Location = new Point(btn_OK.Location.X + btn_OK.Width + 7, btn_OK.Location.Y);
				btn_Cancel.Location = new Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_OK.Location.Y);
			}
			else
			{
				this.Text = UpdatedConversionParameters.PostProcessSettings.NiceName;
				panel1.Visible = false;
				oLayoutPanel.Visible = false;
				mLayoutPanel.Location = new Point(panel1.Location.X, panel1.Location.Y);
				grpBox_Properties.Location = new Point(mLayoutPanel.Location.X, mLayoutPanel.Location.Y + 30);

				this.Height = this.Height - panel1.Height - oLayoutPanel.Height - btn_HideDetail.Height - 130 ;
				int tabIndex = 0;
				int w = 0;
				int h = 0;

				//shaby (begin): Implementing TableLayoutPannel
				int oTableLayoutCurrentRow = 0;
				int oTableLayoutCurrentColumn = 0;
				oTableLayoutPannel.Visible = false;
				oTableLayoutPannel.Name = "oTableLayoutPannel";
				oTableLayoutPannel.AutoSize = true;
				oTableLayoutPannel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
				//oTableLayoutPannel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
				oTableLayoutPannel.TabIndex = btn_OpenDetail.TabIndex + 1;
				oTableLayoutPannel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
				//shaby (end)  : Implementing TableLayoutPannel

				foreach (ScriptParameter p in UpdatedConversionParameters.PostProcessSettings.ParameterList)
				{
					PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();

					strBrtextBox = prepopulateDaisyOutput != null
									   ? prepopulateDaisyOutput.OutputPath
									   : string.Empty;

					if (p.Name != "input" && p.ParameterDataType is PathDataType && p.IsParameterRequired)
					{
						Control c = (Control)new PathBrowserControl(p, inputFileName, strBrtextBox);
						c.Anchor = AnchorStyles.Right;
						mLayoutPanel.Controls.Add(c);
						mLayoutPanel.SetFlowBreak(c, true);
						c.TabIndex = tabIndex++;
						if (w < c.Width + c.Margin.Horizontal) w = c.Width + c.Margin.Horizontal;
						h += c.Height + c.Margin.Vertical;

					}
					else if (!p.IsParameterRequired)
					{
						//TODO: added to work epub not. Should add correct procesing of this parameter. 
						if (p.ParameterDataType is PathDataType)
							continue;

						Control c =
							p.ParameterDataType is BoolDataType ? new BoolControl(p)
								: p.ParameterDataType is EnumDataType ? new EnumControl(p)
								: p.ParameterDataType is StringDataType ? new StrUserControl(p)
								: (Control)new IntUserControl(p);

						c.Anchor = AnchorStyles.Right;
						oTableLayoutPannel.Controls.Add(c, oTableLayoutCurrentColumn, oTableLayoutCurrentRow);
						this.Controls.Add(oTableLayoutPannel);
						oTableLayoutCurrentRow++;

					}
				}
				oTableLayoutPannel.Location = new Point(btn_OpenDetail.Location.X, btn_OpenDetail.Location.Y + 5);

				if (oTableLayoutPannel.Width > mLayoutPanel.Width)
				{
					mLayoutPanel.Controls[0].Width = oTableLayoutPannel.Width - 3;
					tableLayoutPanel2.Width = oTableLayoutPannel.Width - 4;

				}
				else
				{
					oTableLayoutPannel.Controls[0].Width = mLayoutPanel.Width - 3;
					tableLayoutPanel2.Width = mLayoutPanel.Width - 4;
				}

				btn_OK.Location = new Point(oTableLayoutPannel.Width - 220, btn_OpenDetail.Location.Y);
				btn_Reset.Location = new Point(btn_OK.Location.X + btn_OK.Width + 7, btn_OpenDetail.Location.Y);
				btn_Cancel.Location = new Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_OpenDetail.Location.Y);
			}
		}

		/// <summary>
		/// Function to show the Browse for Folder Dialog box
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btn_Browse_Click(object sender, EventArgs e)
		{
			String path = GetPhysicalPath();
			if (path != "")
				tBx_Browse.Text = path;
		}

        private void AccessibilityCheckerLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
			Process.Start("https://support.microsoft.com/en-us/office/improve-accessibility-with-the-accessibility-checker-a16f6de0-2f39-4a2b-8bd8-5ad801426c7f#bkmk_use");
		}

       


        /// <summary>
        /// Function to reset all the fields in the UI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Reset_Click(object sender, EventArgs e)
		{
			tBx_Browse.Text = "";
			tBx_Title.Text = "";
			tBx_Creator.Text = "";
			tBx_Publisher.Text = "";
			tBx_Uid.Text = "";
			int counter = 0;
			if (ConverterHelper.PipelineIsInstalled() && useAScript)
			{
				mLayoutPanel.Controls[0].Controls[0].Controls[1].Text = "";
				foreach (Control c in oTableLayoutPannel.Controls)
				{
					if (oTableLayoutPannel.Controls[counter].Name == "EnumControl")
					{
						ComboBox comboBox = (ComboBox)oTableLayoutPannel.Controls[counter].Controls[1];
						comboBox.SelectedIndex = 0;
					}
					else if (oTableLayoutPannel.Controls[counter].Name == "BoolControl")
					{
						CheckBox checkBox = (CheckBox)oTableLayoutPannel.Controls[counter].Controls[0];
						checkBox.Checked = false;
					}
					counter++;
				}
			}
		}



		/// <summary>
		/// Function to Continue the Translation process
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public void btn_OK_Click(object sender, EventArgs e)
		{
			try
			{
				Translate();
				
			}
			catch (Exception ex)
			{
				AddinLogger.Error(ex);
				MessageBox.Show(ex.Message);
			}
		}

		private void btn_HideDetail_Click(object sender, EventArgs e)
		{

			btn_OpenDetail.Visible = true;
			btn_HideDetail.Visible = false;
			oTableLayoutPannel.Visible = false;

			//panelButton.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 220, btn_OpenDetail.Location.Y - 3);
			btn_OK.Location = new Point(oTableLayoutPannel.Width - 220, btn_OpenDetail.Location.Y);
			btn_Reset.Location = new Point(btn_OK.Location.X + btn_OK.Width + 7, btn_OpenDetail.Location.Y);
			btn_Cancel.Location = new Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_OpenDetail.Location.Y);
		}

		private void btn_Cancel_Click(object sender, EventArgs e)
		{
			string[] files = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY\");
			foreach (string file in files)
			{
				if (file.Contains(".PNG") || file.Contains(".png"))
				{
					File.Delete(file);
				}
			}
			this.Close();
		}

		private void btn_OpenDetail_Click_1(object sender, EventArgs e)
		{
			btn_OpenDetail.Visible = false;
			btn_HideDetail.Visible = true;
			oTableLayoutPannel.Visible = true;

			//dynamically setting tab index for the usercontrols based on the order it is placed
			int oTableLayoutPannelTabIndex = oTableLayoutPannel.TabIndex;
			foreach (Control ctrl in oTableLayoutPannel.Controls)
			{
				if (ctrl.Name == "EnumControl")
					ctrl.Controls[1].TabIndex = ++oTableLayoutPannelTabIndex;
				else if (ctrl.Name == "BoolControl")
					ctrl.Controls[0].TabIndex = ++oTableLayoutPannelTabIndex;
			}
			oTableLayoutPannel.Controls[0].Focus();


			btn_HideDetail.Location = new Point(oTableLayoutPannel.Location.X, oTableLayoutPannel.Location.Y + oTableLayoutPannel.Height + 2);

			//panelButton.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 220, btn_HideDetail.Location.Y -3);
			btn_OK.Location = new Point(oTableLayoutPannel.Width - 220, btn_HideDetail.Location.Y);
			btn_Reset.Location = new Point(btn_OK.Location.X + btn_OK.Width + 7, btn_HideDetail.Location.Y);
			btn_Cancel.Location = new Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_HideDetail.Location.Y);
		}

		#endregion
	}
}