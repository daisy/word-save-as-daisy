using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Xml;
using System.Resources;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.IO.Packaging;
using System.Drawing.Imaging;
using Daisy.SaveAsDAISY.Forms;
using Daisy.SaveAsDAISY.Forms.Controls;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class MultipleSub : Form
    {
        const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        PackageRelationship packRelationship = null;
        const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        List<string> listDocuments;
        private ResourceManager Labels;
        Hashtable table = new Hashtable();
        int subCount = 0;
        int masterSubFlag = 0;
        string fileOutputPath, versionInfo = "";
        private ScriptParser mParser = null;
        private bool useAScript = false;
        String input = "", uId = "";
        private string mProjectDirectory;
        string strBrtextBox = "";
        TableLayoutPanel oTableLayoutPannel = new TableLayoutPanel();
       

        public ConversionParameters UpdatedConversionParameters { get; private set; }

        public MultipleSub(ConversionParameters conversion) {
            UpdatedConversionParameters = conversion;
            InitializeComponent();
            listDocuments = new List<string>();
            this.Labels = AddInHelper.LabelsManager;
            btn_Up.Enabled = false;
            btn_Down.Enabled = false;
            btn_Delete.Enabled = false;
            btn_Populate.Enabled = false;
            this.versionInfo = UpdatedConversionParameters.Version;
            useAScript = UpdatedConversionParameters.ScriptPath != null && UpdatedConversionParameters.ScriptPath.Length > 0;
            if (useAScript) {
                if(UpdatedConversionParameters.PostProcessSettings == null )
                    UpdatedConversionParameters.PostProcessSettings = new ScriptParser(UpdatedConversionParameters.ScriptPath);
                
                this.Text = UpdatedConversionParameters.PostProcessSettings.NiceName;
            }
            
        }
        /// <summary>
        /// Prpoperty to return Selected Documents
        /// </summary>
        public List<string> GetFileNames
        {
            get { return listDocuments; }
        }
        /// <summary>
        /// Prpoperty to return Output file path
        /// </summary>
        public string GetOutputFilePath
        {
            get { return fileOutputPath; }
        }

        public string pipeOutput
        {
            get { return strBrtextBox; }
        }

        /* Returns a Hash Table having information about Title,Creator,Publisher,UID*/
        public Hashtable HTable
        {
            get { return table; }
        }

        /* Returns false when Form is diplayed to the user*/
        public int DoTranslate()
        {
            this.ShowDialog();
            return masterSubFlag;
        }

        /* User can take select Output folder and function returns the Output filename chosen by the user*/
        private string GetPhysicalPath()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Destination";
            if (tBx_Browse.Text != "" && Directory.Exists(tBx_Browse.Text))
                fbd.SelectedPath = tBx_Browse.Text;
            fbd.ShowDialog();
            return fbd.SelectedPath;
            // return the path in which the user wants to create the file
        }

        /*Core function for Translation*/
        private void btn_Ok_Click(object sender, EventArgs e)
        {
        	CleanOutputDirDlg cleanUpDialog;

            if (!useAScript) {
                if (lBx_SubDocs.Items.Count == 0) {
                    MessageBox.Show(Labels.GetString("SubdocsError"), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_Browse.Focus();
                } else if (tBx_output.Text == "") {
                    MessageBox.Show(Labels.GetString("ChoseDestinationFile"), Labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_Output.Focus();
                } else if (Directory.Exists(Path.GetDirectoryName(tBx_output.Text)) == false) {
                    MessageBox.Show("Directory " + string.Concat(Path.GetDirectoryName(tBx_output.Text), " ", "does not exist"), Labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_Output.Focus();
                } else if (Path.GetFileNameWithoutExtension(tBx_output.Text) == "") {
                    MessageBox.Show("Please provide proper filename", Labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_Output.Focus();
                } else if (tBx_Title.Text.TrimEnd() == "") {
                    MessageBox.Show("Please enter the Title", Labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tBx_Title.Focus();
                } else {
                    for (int i = 0; i < lBx_SubDocs.Items.Count; i++)
                    {
                        listDocuments.Add(lBx_SubDocs.Items[i].ToString());
                        btn_Delete.Enabled = true;
                        btn_Populate.Enabled = true;
                    }

                    UpdatedConversionParameters = UpdatedConversionParameters.withParameter("Title", tBx_Title.Text)
                        .withParameter("Creator", tBx_Creator.Text)
                        .withParameter("Publisher", tBx_Publisher.Text)
                        .withParameter("Subject", tBx_Subject.Text)
                        .withParameter("MasterSub", "Yes");


                    if (tBx_Uid.Text != "")
                    {
                        if (uId == tBx_Uid.Text)
                            UpdatedConversionParameters.withParameter("UID", "AUTO-UID-" + tBx_Uid.Text);
                        else
                            UpdatedConversionParameters.withParameter("UID", tBx_Uid.Text);
                    }
                    else
                        UpdatedConversionParameters.withParameter("UID", "AUTO-UID-" + GenerateId().ToString());


                    if (Path.GetExtension(tBx_output.Text) == "")

                        tBx_output.Text = tBx_output.Text + ".xml";

                    fileOutputPath = tBx_output.Text;
                    UpdatedConversionParameters.OutputPath = fileOutputPath;
                    // Now retrieved by the conversion parameters class while building the params hash
                    /*ConverterSettings daisySt = new ConverterSettings();
                    String imgoption = daisySt.GetImageOption;
                    String resampleValue = daisySt.GetResampleValue;
                    String characterStyle = daisySt.GetCharacterStyle;
                    String pagenumStyle = daisySt.GetPagenumStyle;
                    if (imgoption != " ")
                    {
                        table.Add("ImageSizeOption", imgoption);
                        table.Add("DPI", resampleValue);
                    }
                    if (characterStyle != " ")
                    {
                        table.Add("CharacterStyles", characterStyle);
                    }
                    if (pagenumStyle != " ")
                    {
                        table.Add("Custom", pagenumStyle);
                    }*/

                    masterSubFlag = 1;
                    this.Close();
                }
            } else {
            	string scriptOutput = string.Empty;
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
								if (!string.IsNullOrEmpty(pathDataType.FileExtenssion) &&
								    !pathDataType.FileExtenssion.Equals(outputFileInfo.Extension, StringComparison.InvariantCultureIgnoreCase))
								{
									MessageBox.Show(string.Format("Please select {0} output file", pathDataType.FileExtenssion), "Error",
									                MessageBoxButtons.OK, MessageBoxIcon.Error);
									mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
									return;
								}
								strBrtextBox = outputFileInfo.DirectoryName;
								scriptOutput = outputFileInfo.Name;
							}
							catch (ArgumentException ex)
							{
								AddinLogger.Error(ex);
								MessageBox.Show("Please select output file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
								mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
								return;
							}

						}
						else
						{
							strBrtextBox = p.ParameterValue;
						}
						break;

					}
				}

            	cleanUpDialog = new CleanOutputDirDlg(strBrtextBox, scriptOutput);

            	if (lBx_SubDocs.Items.Count == 0) {
                    MessageBox.Show(Labels.GetString("SubdocsError"), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btn_Browse.Focus();
                } else if (tBx_Title.Text.TrimEnd() == "") {
                    MessageBox.Show("Please enter the Title", Labels.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tBx_Title.Focus();
                } else if (strBrtextBox.TrimEnd() == "") {
                    MessageBox.Show("Please select the Destination folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
                } else if (cleanUpDialog.Clean(this) == DialogResult.Cancel) {
                    mLayoutPanel.Controls[0].Controls[0].Controls[1].Focus();
                } else {
					if (strBrtextBox != cleanUpDialog.OutputDir)
					{
						strBrtextBox = cleanUpDialog.OutputDir;
						foreach (ScriptParameter p in UpdatedConversionParameters.PostProcessSettings.ParameterList)
						{
							if (p.IsParameterRequired && (p.Name == "outputPath" || p.Name == "output"))
							{
								p.ParameterValue = cleanUpDialog.OutputDir;
							}
						}
					}


					tBx_output.Text = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY";
                    for (int i = 0; i < lBx_SubDocs.Items.Count; i++)
                    {
                        listDocuments.Add(lBx_SubDocs.Items[i].ToString());
                        btn_Delete.Enabled = true;
                        btn_Populate.Enabled = true;
                    }

                    UpdatedConversionParameters.withParameter("Title", tBx_Title.Text)
                        .withParameter("Creator", tBx_Creator.Text)
                        .withParameter("Publisher", tBx_Publisher.Text)
                        .withParameter("Subject", tBx_Subject.Text)
                        .withParameter("MasterSub", "Yes")
                        .withParameter("PipelineOutput", strBrtextBox);

                    if (tBx_Uid.Text != "")
                        if (uId == tBx_Uid.Text)
                            UpdatedConversionParameters.withParameter("UID", "AUTO-UID-" + tBx_Uid.Text);
                        else
                            UpdatedConversionParameters.withParameter("UID", tBx_Uid.Text);
                    else
                        UpdatedConversionParameters.withParameter("UID", "AUTO-UID-" + GenerateId().ToString());

                    if (Path.GetExtension(tBx_output.Text) == "")
                        tBx_output.Text = Path.Combine(tBx_output.Text, "MultipleNarrator" + ".xml");

                    fileOutputPath = tBx_output.Text;
                    UpdatedConversionParameters.OutputPath = fileOutputPath;

                    masterSubFlag = 1;
                    this.Close();
                }
            }

        }

        /*Function to Close UI*/
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

        /*Function to select Output folder*/
        private void btn_Browse_Click(object sender, EventArgs e)
        {
            //Browse button click
            String path = GetPhysicalPath();
            if (path != "")
            {
                tBx_Browse.Text = path;
                SearchOption option = SearchOption.TopDirectoryOnly;
                string[] files = Directory.GetFiles(path, "*.docx", option);
                foreach (string input in files)
                {
                    if (!Path.GetFileNameWithoutExtension(input).StartsWith("~$") && input.EndsWith(".docx"))
                    {
                        lBx_SubDocs.Items.Insert(subCount, input);
                        subCount++;
                    }
                }

                if (lBx_SubDocs.Items.Count > 0)
                {
                    lBx_SubDocs.SetSelected(0, true);
                    //btn_Up.Enabled = true;
                    btn_Down.Enabled = true;
                    btn_Delete.Enabled = true;
                    btn_Populate.Enabled = true;
                }
            }
        }

        /* User can take select Output folder and function returns the Output filename chosen by the user*/
        private string GetPhysicalPath_Output()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "XML files (*.xml)|*.xml";
            sfd.FilterIndex = 1;
            sfd.CheckPathExists = true;
            sfd.DefaultExt = ".xml";
            sfd.ShowDialog();
            return sfd.FileName;
            // return the path in which the user wants to create the file
        }

        /*Function to select Output File*/
        private void btn_Output_Click(object sender, EventArgs e)
        {
            //Output buttonClick
            String path = GetPhysicalPath_Output();
            fileOutputPath = path;
            if (path != "")

                tBx_output.Text = path;

        }

        /// <summary>
        /// Function to change the position of the Document 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Up_Click(object sender, EventArgs e)
        {
            //Move up button click
            if (lBx_SubDocs.SelectedItem.ToString() != "")
            {
                int currentIndx = lBx_SubDocs.SelectedIndex;
                int upIndx = currentIndx - 1;

                String currentIndx_Value = lBx_SubDocs.Items[currentIndx].ToString();
                String upIndx_Value = lBx_SubDocs.Items[upIndx].ToString();

                lBx_SubDocs.Items[currentIndx] = upIndx_Value;
                lBx_SubDocs.Items[upIndx] = currentIndx_Value;
                lBx_SubDocs.SetSelected(upIndx, true);
            }
        }

        /// <summary>
        /// Function to change the position of the Document 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Down_Click(object sender, EventArgs e)
        {
            //Move down button click
            if (lBx_SubDocs.SelectedItem.ToString() != "")
            {
                int currentIndx = lBx_SubDocs.SelectedIndex;
                int downIndx = currentIndx + 1;

                String currentIndx_Value = lBx_SubDocs.Items[currentIndx].ToString();
                String downIndx_Value = lBx_SubDocs.Items[downIndx].ToString();

                lBx_SubDocs.Items[currentIndx] = downIndx_Value;
                lBx_SubDocs.Items[downIndx] = currentIndx_Value;
                lBx_SubDocs.SetSelected(downIndx, true);
            }
        }

        /// <summary>
        /// Function to remove the Document 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Delete_Click(object sender, EventArgs e)
        {
            //Remove button click
            int currentIndx = lBx_SubDocs.SelectedIndex;
            if (currentIndx >= 0)
            {
                lBx_SubDocs.Items.RemoveAt(currentIndx);
                subCount--;
            }
            lBx_SubDocs.Refresh();

            if (lBx_SubDocs.Items.Count <= 0)
            {
                btn_Delete.Enabled = false;
                btn_Populate.Enabled = false;
            }
            else
                lBx_SubDocs.SetSelected(0, true);
        }

        /*Function to reset all the values in UI*/
        private void btn_Reset_Click(object sender, EventArgs e)
        {
            tBx_Browse.Text = "";
            tBx_output.Text = "";
            tBx_Creator.Text = "";
            tBx_Title.Text = "";
            tBx_Subject.Text = "";
            tBx_Publisher.Text = "";
            tBx_Uid.Text = "";
            lBx_SubDocs.Items.Clear();
            btn_Up.Enabled = false;
            btn_Down.Enabled = false;
            btn_Delete.Enabled = false;
            btn_Populate.Enabled = false;
            subCount = 0;
            int counter = 0;
            if (useAScript )
            //if (useAScript) 
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

            //mLayoutPanel.Controls[0].Controls[0].Text = "";
        }

        /*Function to select a document*/
        private void button1_Click(object sender, EventArgs e)
        {
            //Add button click
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Word Document(*.docx)|*.docx";
            ofd.FilterIndex = 1;
            ofd.CheckPathExists = true;
            ofd.DefaultExt = ".docx";
            ofd.ShowDialog();
            if (ofd.FileName != "")
            {
                lBx_SubDocs.Items.Insert(subCount, ofd.FileName);
                subCount++;
            }
            lBx_SubDocs.Refresh();

            if (lBx_SubDocs.Items.Count > 0)
            {
                lBx_SubDocs.SetSelected(0, true);
                btn_Delete.Enabled = true;
                btn_Populate.Enabled = true;
            }

        }

        /*Function to load form*/
        private void MultipleSub_Load(object sender, EventArgs e)
        {
            tBx_Uid.Text = GenerateId().ToString();
            uId = tBx_Uid.Text;
            if (!useAScript)
            {
                mLayoutPanel.Visible = false;
                oLayoutPanel.Visible = false;
                grpBox_Properties.Location = mLayoutPanel.Location;
                btnShow.Visible = false;
                btnHide.Visible = false;
                oLayoutPanel.Visible = false;

                btn_Ok.Location = new System.Drawing.Point(grpBox_Properties.Width - 220, grpBox_Properties.Location.Y + grpBox_Properties.Height + 3);
                btn_Reset.Location = new System.Drawing.Point(btn_Ok.Location.X + btn_Ok.Width + 7, btn_Ok.Location.Y);
                btn_Cancel.Location = new System.Drawing.Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_Ok.Location.Y);
            }
            else
            {
                this.Text = UpdatedConversionParameters.PostProcessSettings.NiceName;
                groupBoxXml.Visible = false;
                btnShow.Visible = true;
                btnHide.Visible = false;
                oLayoutPanel.Visible = false;


                int destinationInitialtabIndex = btn_Down.TabIndex + 1;
                int w = 0;
                int h = 0;

                //shaby (begin): Implementing TableLayoutPannel
                int oTableLayoutCurrentRow = 0;
                int oTableLayoutCurrentColumn = 0;
                oTableLayoutPannel.Visible = false;
                oTableLayoutPannel.Name = "oTableLayoutPannel";
                oTableLayoutPannel.AutoSize = true;
                oTableLayoutPannel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                //oTableLayoutPannel.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
                oTableLayoutPannel.TabIndex = btnShow.TabIndex + 1; //60 - 100 (reserved for advanced user controls)
                oTableLayoutPannel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
                //shaby (end)  : Implementing TableLayoutPannel

                foreach (ScriptParameter p in UpdatedConversionParameters.PostProcessSettings.ParameterList)
                {
                    if (p.Name != "input" && p.ParameterDataType is PathDataType && p.IsParameterRequired)
                    {
                        Control c = (Control)new PathBrowserControl(p, input, mProjectDirectory);
                        c.Anchor = AnchorStyles.Right;
                        c.Anchor = AnchorStyles.Top;
                        mLayoutPanel.Controls.Add(c);
                        mLayoutPanel.SetFlowBreak(c, true);
                        c.TabIndex = destinationInitialtabIndex++; //reserved from 9 to 50
                        if (w < c.Width + c.Margin.Horizontal) w = c.Width + c.Margin.Horizontal + 30;
                        h += c.Height + c.Margin.Vertical;
                    }
                    else if (!p.IsParameterRequired)
                    {
						//TODO: 
						if (p.ParameterDataType is PathDataType)
							continue;

                        Control c =
                            p.ParameterDataType is BoolDataType ? (Control)new BoolControl(p) :
                            p.ParameterDataType is EnumDataType ? (Control)new EnumControl(p) :
                            p.ParameterDataType is StringDataType ? (Control)new StrUserControl(p) :
                            (Control)new IntUserControl(p);
                        c.Anchor = AnchorStyles.Right;
                        oTableLayoutPannel.Controls.Add(c, oTableLayoutCurrentColumn, oTableLayoutCurrentRow);
                        this.Controls.Add(oTableLayoutPannel);
                        oTableLayoutCurrentRow++;
                    }
                }

                mLayoutPanel.Location = new System.Drawing.Point(groupBoxXml.Location.X, groupBoxXml.Location.Y);
                grpBox_Properties.Location = new System.Drawing.Point(mLayoutPanel.Location.X, mLayoutPanel.Location.Y + mLayoutPanel.Height);
                btnShow.Location = new System.Drawing.Point(grpBox_Properties.Location.X, grpBox_Properties.Location.Y + grpBox_Properties.Height);
                oTableLayoutPannel.Location = new System.Drawing.Point(btnShow.Location.X, btnShow.Location.Y);
                //btnHide.Location = new System.Drawing.Point(oLayoutPanel.Location.X, oLayoutPanel.Location.Y + 157);

                oTableLayoutPannel.Location = new System.Drawing.Point(btnShow.Location.X, btnShow.Location.Y);
                if (oTableLayoutPannel.Width > mLayoutPanel.Width)
                {
                    mLayoutPanel.Controls[0].Width = oTableLayoutPannel.Width - 3;
                    tableLayoutPanel1.Width = oTableLayoutPannel.Width - 4;
                    tableLayoutPanel2.Width = oTableLayoutPannel.Width;
                }
                else
                {
                    oTableLayoutPannel.Controls[0].Width = mLayoutPanel.Width - 3;
                    tableLayoutPanel1.Width = mLayoutPanel.Width - 4;
                    tableLayoutPanel2.Width = mLayoutPanel.Width;
                }
                btn_Ok.Location = new System.Drawing.Point(grpBox_Properties.Width - 232, grpBox_Properties.Location.Y + grpBox_Properties.Height + 3);
                btn_Reset.Location = new System.Drawing.Point(btn_Ok.Location.X + btn_Ok.Width + 7, btn_Ok.Location.Y);
                btn_Cancel.Location = new System.Drawing.Point(btn_Reset.Location.X + btn_Reset.Width + 7, btn_Ok.Location.Y);
            }
        }

        /* Function to Generate an unique ID */
        public long GenerateId()
        {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        /*Core Function to populate the document properties of the selcted document*/
        private void btn_Populate_Click(object sender, EventArgs e)
        {
            if (tBx_Title.Text != "" || tBx_Creator.Text != "" || tBx_Publisher.Text != "" || tBx_Subject.Text != "")
            {
                DialogResult dr = MessageBox.Show("Do you want to overwrite the document properties?", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                {
                    populateProperties();
                }
            }
            else
                populateProperties();

        }

        /*Function to populate the document properties of the selcted document*/
        public void populateProperties()
        {
            String fileSelected = lBx_SubDocs.SelectedItem.ToString();

            if (fileSelected != "")
            {
                String resultOpenSub = CheckFileOPen(fileSelected);

                if (resultOpenSub == "notopen")
                {
                    Package pack = Package.Open(fileSelected, FileMode.Open, FileAccess.ReadWrite);

                    tBx_Creator.Text = pack.PackageProperties.Creator;
                    tBx_Subject.Text = pack.PackageProperties.Subject;
                    pack.Close();
                    tBx_Title.Text = DocPropTitle(fileSelected);
                    tBx_Publisher.Text = DocPropPublisher(fileSelected);
                }
                else
                    MessageBox.Show(Labels.GetString("Populateopen"), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public String DocPropPublisher(String docFile)
        {
            Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(appRelationshipType))
            {
                packRelationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            XmlDocument doc = new XmlDocument();
            doc.Load(mainPartxml.GetStream());

            pack.Close();

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("vt", appNamespace);

            XmlNodeList node = doc.GetElementsByTagName("Company");
            if (node != null)
                return node.Item(0).InnerText;
            else
                return "";
        }

        public String DocPropTitle(String docFile)
        {
            int titleFlag = 0;
            String styleVal = "";
            string temp = "";
            Package pack;
            pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
            String title = pack.PackageProperties.Title;
            pack.Close();
            if (title != "")
                return title;
            else
            {
                pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
                {
                    packRelationship = searchRelation;
                    break;
                }
                Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(mainPartxml.GetStream());
                NameTable nt = new NameTable();
                pack.Close();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                XmlNodeList getParagraph = doc.SelectNodes("//w:body/w:p/w:pPr/w:pStyle", nsManager);
                for (int j = 0; j < getParagraph.Count; j++)
                {
                    XmlAttributeCollection paraGraphAttribute = getParagraph[j].Attributes;
                    for (int i = 0; i < paraGraphAttribute.Count; i++)
                    {
                        if (paraGraphAttribute[i].Name == "w:val")
                        {
                            styleVal = paraGraphAttribute[i].Value;
                        }
                        if (styleVal != "" && styleVal == "Title")
                        {
                            XmlNodeList getStyle = getParagraph[j].ParentNode.ParentNode.SelectNodes("w:r", nsManager);
                            if (getStyle != null)
                            {
                                for (int k = 0; k < getStyle.Count; k++)
                                {
                                    XmlNode getText = getStyle[k].SelectSingleNode("w:t", nsManager);
                                    temp = temp + " " + getText.InnerText;
                                }
                            }
                            titleFlag = 1;
                            break;
                        }
                        if (titleFlag == 1)
                        {
                            break;
                        }
                    }
                    if (titleFlag == 1)
                    {
                        break;
                    }
                }

                title = temp;

            }
            return title;
        }

        /*Function to check whether selected document is open or not*/
        public String CheckFileOPen(String listSubDocs)
        {
        	AddinLogger.Info("Check if file opened : " + listSubDocs);
            String resultSubDoc = "notopen";
            try
            {
                Package pack;
                pack = Package.Open(listSubDocs, FileMode.Open, FileAccess.ReadWrite);
                pack.Close();
            }
            catch (Exception e)
            {
				AddinLogger.Error(e);
                resultSubDoc = "open";
            }
            return resultSubDoc;
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            btnShow.Visible = false;
            btnHide.Visible = true;
            oTableLayoutPannel.Visible = true;

            //dynamically setting tab index for the usercontrols based on the order it is placed
            int oTableLayoutPannelTabIndex = oTableLayoutPannel.TabIndex;
            foreach (Control ctrl in oTableLayoutPannel.Controls)
            {   //reserved tabIndex 61 - 100 for Advanced User-controls
                if (ctrl.Name == "EnumControl")
                    ctrl.Controls[1].TabIndex = ++oTableLayoutPannelTabIndex;
                else if (ctrl.Name == "BoolControl")
                    ctrl.Controls[0].TabIndex = ++oTableLayoutPannelTabIndex;
            }
            oTableLayoutPannel.Controls[0].Focus();


            btnHide.Location = new System.Drawing.Point(oTableLayoutPannel.Location.X, oTableLayoutPannel.Location.Y + oTableLayoutPannel.Height + 2);
            //panelButton.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 220, btnHide.Location.Y - 3);
            btn_Ok.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 225, btnHide.Location.Y);
            btn_Reset.Location = new System.Drawing.Point(btn_Ok.Location.X + btn_Ok.Width + 7, btnHide.Location.Y);
            btn_Cancel.Location = new System.Drawing.Point(btn_Reset.Location.X + btn_Reset.Width + 7, btnHide.Location.Y);
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            btnShow.Visible = true;
            btnHide.Visible = false;
            oTableLayoutPannel.Visible = false;
            //panelButton.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 220, btnShow.Location.Y - 3);
            btn_Ok.Location = new System.Drawing.Point(oTableLayoutPannel.Width - 225, btnShow.Location.Y);
            btn_Reset.Location = new System.Drawing.Point(btn_Ok.Location.X + btn_Ok.Width + 7, btnShow.Location.Y);
            btn_Cancel.Location = new System.Drawing.Point(btn_Reset.Location.X + btn_Reset.Width + 7, btnShow.Location.Y);
        }

        /*Function to update the changes in the document positions*/
        private void lBx_SubDocs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lBx_SubDocs.SelectedIndex == 0)
                btn_Up.Enabled = false;
            else
                btn_Up.Enabled = true;

            if (lBx_SubDocs.SelectedIndex == lBx_SubDocs.Items.Count - 1)
                btn_Down.Enabled = false;
            else
                btn_Down.Enabled = true;

            if (lBx_SubDocs.SelectedIndex >= 0)
            {
                btn_Delete.Enabled = true;
                btn_Populate.Enabled = true;
            }
            else
            {
                btn_Delete.Enabled = false;
                btn_Populate.Enabled = false;
                btn_Up.Enabled = false;
                btn_Down.Enabled = false;
            }

            if (lBx_SubDocs.Items.Count > 0)
            {
                btn_Delete.Enabled = true;
                btn_Populate.Enabled = true;

            }
            else
            {
                btn_Delete.Enabled = false;
                btn_Populate.Enabled = false;
                btn_Up.Enabled = false;
                btn_Down.Enabled = false;
            }
        }


    }
}