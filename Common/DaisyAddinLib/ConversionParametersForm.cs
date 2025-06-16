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
using Daisy.SaveAsDAISY.Forms.Controls;
using System.Globalization;
using System.Diagnostics.Eventing.Reader;

namespace Daisy.SaveAsDAISY.Conversion
{
	/**
	 * Form to graphicaly complete a conversion parameters set
	 */
	public partial class ConversionParametersForm : Form
	{
		int translateFlag = 0;
		XmlDocument XmlPackage;
		Hashtable parametersHash = new Hashtable();
		string officeVersion = "";
		PackageRelationship packRelationship = null;
		string inputFileName, embedFilePath, tempInput, trackChangeFlag, outputFilePath, masterSubFlag;
		const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
		const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		string uId = "";
		string btnID = "";
		private string mInputPath;
		private bool useAScript = false;
		string outputFileOrFolder = "";
		TableLayoutPanel oTableLayoutPannel = new TableLayoutPanel();

		public ConversionParameters UpdatedConversionParameters { get; private set; }

        BaseUserControl convertScriptParameterToControl(ScriptParameter param)
        {
            Type paramType = param.ParameterData.GetType();
            if (paramType == typeof(BoolData))
            {
                return new BoolControl(param);
            }
            else if (
				paramType == typeof(EnumData))
            {
                return new EnumControl(param);
            }
            else if (paramType == typeof(IntegerData))
            {
                return new IntUserControl(param);
            }
            else if (paramType == typeof(PathData))
            {
                return new PathControl(param);
            }
            else if (paramType == typeof(StringData))
            {
                return new StrUserControl(param);
            }
            return null;
        }


        /// <summary>
        /// Returns Hash Table having information about Title,Creator,Publisher,UID
        /// </summary>
        public Hashtable ParametersHash { get { return parametersHash; } }

		/// <summary>
		/// Return Title information 
		/// </summary>
		public string GetTitle { get { return TitleInput.Text; } }

		public Script getParser { get { return UpdatedConversionParameters.PostProcessor; } }

		/// <summary>
		/// Return Creator information 
		/// </summary>
		public string GetCreator { get { return CreatorInput.Text; } }

		/// <summary>
		/// Returns the value whether to translate the current document
		/// along with sub documents
		/// </summary>
		public string MastersubValue { get { return masterSubFlag; } }


		/// <summary>
		/// Returns the path of the output file
		/// </summary>
		public string OutputFilepath { get { return outputFilePath; } }

		public string PipeOutput { get { return outputFileOrFolder; } }


        /// <summary>
        /// Default form for converting a word file
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        public ConversionParametersForm(DocumentParameters document, ConversionParameters conversion) {
            // Copy current conversion settings
            UpdatedConversionParameters = conversion.usingMainDocument(document);
            
            btnID = conversion.ControlName;
			tempInput = document.CopyPath ?? document.InputPath;

			// if a script is defined in the parameters
			useAScript = UpdatedConversionParameters.PostProcessor != null;
			if (useAScript) {
				mInputPath = document.InputPath;

			} else {
				inputFileName = document.InputPath;
			}
			
			InitializeComponent();
			// Loading form content
            UIDTextBox.Text = GenerateId().ToString();
            uId = UIDTextBox.Text;

            CreatorInput.Text = UpdatedConversionParameters.Creator;
            TitleInput.Text = UpdatedConversionParameters.Title;
            PublisherInput.Text = UpdatedConversionParameters.Publisher;

            // Unknown / no language selection
            languageSelector.Items.Add("");
            // first load languages found in document
            if (document.Languages.Count > 0) {
                foreach (string item in document.Languages) {
                    languageSelector.Items.Add(item);
                }
                // select the main most present language
                languageSelector.SelectedIndex = 1;
            }
            // load the remaining list of available language codes
            CultureInfo[] cultures = CultureInfo.GetCultures(CultureTypes.AllCultures);
            foreach (CultureInfo culture in cultures) {
                if (!languageSelector.Items.Contains(culture.Name)) {
                    languageSelector.Items.Add(culture.Name);
                }
            }

            DestinationControl.Width = this.Width - 10;

            //Point temp = AdvancedSettingsPanel.Location;
            //BottomPanel.Location = temp;
            //this.Height = this.Height - AdvancedSettingsPanel.Height;
            AdvancedSettingsGroup.Visible = false;
            //SwitchAdvancedSettingsButton.Visible = false;
            if (!useAScript) {
                AdvancedSettingsGroup.Visible = false;
                PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();

                DestinationControl.SelectedPath = prepopulateDaisyOutput != null
                                      ? prepopulateDaisyOutput.OutputPath
                                      : Path.GetDirectoryName(inputFileName);

            } else {
                //SwitchAdvancedSettingsButton.Visible = false;
                this.Text = UpdatedConversionParameters.PostProcessor.NiceName;

                // Link the original destination selection to the output script parameter
                PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();
                ScriptParameter outputParameter = UpdatedConversionParameters.PostProcessor.Parameters["output"];
                outputParameter.Value = prepopulateDaisyOutput != null
                                       ? prepopulateDaisyOutput.OutputPath
                                       : string.Empty;
                DestinationControl.setLinkedParameter(outputParameter);


                // Build AdvancedSettingsPanel
                AdvancedSettingsGroup.Visible = true;
                int tabIndex = AdvancedSettingsPanel.TabIndex;
                int h = 15; // Starting height
                outputFileOrFolder = prepopulateDaisyOutput != null
                                       ? prepopulateDaisyOutput.OutputPath
                                       : string.Empty;

                foreach (var kv in UpdatedConversionParameters.PostProcessor.Parameters) {
                    ScriptParameter p = kv.Value;

                    // output is put in the dedicated output panel
                    if (kv.Key == "output" || kv.Key == "input" || p.Name == "input") continue;

                    if (p.IsParameterRequired || p.IsParameterDisplayed) {
                        Control c = convertScriptParameterToControl(p);
                        AdvancedSettingsPanel.Controls.Add(c);
                        c.Location = new Point(10, h);
                        c.Width = AdvancedSettingsPanel.Width - 20;
                        c.Anchor = AnchorStyles.Top;
                        //RequiredSettingsPanel.Controls.Add(c);
                        c.TabIndex = tabIndex++;
                        h += c.Height;
                    }
                }
                AdvancedSettingsPanel.Size = new Size(AdvancedSettingsPanel.Size.Width, h + 5);
				BottomPanel.TabIndex = tabIndex++;
				OKButton.TabIndex = tabIndex++;
                ResetButton.TabIndex = tabIndex++;
                CancelButton.TabIndex = tabIndex++;

            }
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


		public void Translate()
		{
			bool isValidInputs = !useAScript
							? ValidateForSingleDaisyTranslate()
							: ValidateForFullDaisyTranslate();

			if (isValidInputs)
			{
				if (!useAScript) outputFilePath = DestinationControl.SelectedPath;

				UpdatePopulateOutputXml(DestinationControl.SelectedPath);

				parametersHash = updateParametersHash();

				translateFlag = 1;
				this.Close();
			}

		}

		private bool ValidateForFullDaisyTranslate()
		{
			
			string fileName = Path.GetFileNameWithoutExtension(mInputPath);
			string otpfileName = fileName + ".xml";
            DestinationControl.UpdateScriptParameterValue();
            //DestinationControl.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY\";
			if (TitleInput.Text.TrimEnd() == "")
			{
				MessageBox.Show(Labels.EnterTitle,Labels.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
				TitleInput.Focus();
				return false;
			}


            for (int i = 0; i < AdvancedSettingsGroup.Controls.Count; i++)
			{
				if (AdvancedSettingsGroup.Controls[i] is BaseUserControl)
				{

					((BaseUserControl)AdvancedSettingsGroup.Controls[i]).UpdateScriptParameterValue();

				}
			}


			foreach (var kv in UpdatedConversionParameters.PostProcessor.Parameters)
			{
				var p = kv.Value;
				if (p.IsParameterRequired && (kv.Key == "outputPath" || kv.Key == "output"))
				{
					PathData pathDataType = p.ParameterData as PathData;
					if (pathDataType == null) continue;

					if (pathDataType.IsFile)
					{
						try
						{
							FileInfo outputFileInfo = new FileInfo(p.Value.ToString());
							if (!string.IsNullOrEmpty(pathDataType.FileExtension) && !pathDataType.FileExtension.Equals(outputFileInfo.Extension, StringComparison.InvariantCultureIgnoreCase))
							{
								MessageBox.Show(string.Format(Labels.SelectTypedOutputFile, pathDataType.FileExtension), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
								//AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
								return false;
							}
							outputFileOrFolder = outputFileInfo.DirectoryName;
						}
                        catch (ArgumentException ex)
                        {
                            AddinLogger.Error(ex);
                            MessageBox.Show("Please select output file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
                            return false;
                        }

                    }
					else
					{
						try
						{
                            DirectoryInfo outputFolderInfo = new DirectoryInfo(p.Value.ToString());
                            //outputFileOrFolder = p.Value.ToString();
                            //if (outputFolderInfo.Exists && outputFolderInfo.GetFiles().Length > 0)
                            //{
                            //    MessageBox.Show(string.Format(Labels.SelectEmptyFolder, outputFolderInfo.FullName), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //    return false;
                            //}
                            outputFileOrFolder = outputFolderInfo.FullName;
                        }
                        catch (ArgumentException ex)
                        {
                            AddinLogger.Error(ex);
                            MessageBox.Show("Please select an output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
                            return false;
                        }

                    }

					//TODO:::
					if (string.IsNullOrEmpty(outputFileOrFolder.TrimEnd()))
					{
						MessageBox.Show(Labels.DaisyOutputFolder, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						//AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
						return false;
					}

					CleanOutputDirDlg cleanOutputDirDlg = new CleanOutputDirDlg(outputFileOrFolder, otpfileName);
					if (cleanOutputDirDlg.Clean(this) == DialogResult.Cancel)
					{
						//AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
						return false;
					}
					if (outputFileOrFolder != cleanOutputDirDlg.OutputDir)
					{
						outputFileOrFolder = cleanOutputDirDlg.OutputDir;
						p.Value = cleanOutputDirDlg.OutputDir;
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

			if (string.IsNullOrEmpty(DestinationControl.SelectedPath.TrimEnd()))
			{
				MessageBox.Show(Labels.ChoseDestination, Labels.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
                DestinationControl.Focus();
				return false;
			}

			if (Directory.Exists(DestinationControl.SelectedPath) == false)
			{
				MessageBox.Show(
					string.Format(Labels.DestinationDoesNotExists, DestinationControl.SelectedPath),
					Labels.Error,
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
				);
				return false;
			}

			if (TitleInput.Text.TrimEnd() == "")
			{
				MessageBox.Show(Labels.EnterTitle, Labels.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);
				TitleInput.Focus();
				return false;
			}

			if (File.Exists(DestinationControl.SelectedPath + "\\" + otpfileName))
			{
				DialogResult objResult;
				objResult = MessageBox.Show(string.Format(Labels.RequestFileReplace, otpfileName), Labels.RequestFileReplaceDialogTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
				if (objResult != DialogResult.Yes)
				{
                    DestinationControl.Focus();
					return false;
				}
			}

			return true;
		}

		private Hashtable updateParametersHash()
		{

			UpdatedConversionParameters = UpdatedConversionParameters
				.withParameter("Language", languageSelector.SelectedItem)
				.withParameter("OutputFile", DestinationControl.SelectedPath)
				.withParameter("Title", TitleInput.Text)
				.withParameter("Creator", CreatorInput.Text)
				.withParameter("Publisher", PublisherInput.Text)
				.withParameter("UID", uId == UIDTextBox.Text ? "AUTO-UID-" + UIDTextBox.Text : UIDTextBox.Text)
				.withParameter("Subject", string.Empty)
				.withParameter("PipelineOutput", PipeOutput);


			return UpdatedConversionParameters.ParametersHash;
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
				DialogResult dr = MessageBox.Show(Labels.TrackConfirmation, "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

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
				DialogResult dr = MessageBox.Show(Labels.MasterSubConfirmation, "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

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
			DestinationControl.SelectedPath = "";
			TitleInput.Text = "";
			CreatorInput.Text = "";
			PublisherInput.Text = "";
			UIDTextBox.Text = "";
			int counter = 0;
			if (useAScript)
			{

				//AdancedSettingsLayoutPanel.Controls[0].Controls[0].Controls[1].Text = "";
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
		public void onClickTranslateButton(object sender, EventArgs e)
		{
			try
			{
				Translate();
			}
			catch (Exception ex)
			{
                MessageBox.Show(ex.Message);
                AddinLogger.Error(ex);
			}
		}

		private void onClickCancelButton(object sender, EventArgs e)
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

		private void OnClickSwitchAdvandedSettingsButton(object sender, EventArgs e)
		{
			
			if(AdvancedSettingsGroup.Visible) // Visible, hide the panel
			{
				SwitchAdvancedSettingsButton.Text = "Show Ad&vanced <<";
				AdvancedSettingsGroup.Visible = false;

            } else // not visible, show the panel
			{
				AdvancedSettingsGroup.Visible = true;
                SwitchAdvancedSettingsButton.Text = "Hide Ad&vanced <<";
            }
		}

		#endregion
	}
}