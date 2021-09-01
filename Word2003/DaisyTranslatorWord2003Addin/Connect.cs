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

namespace Daisy.SaveAsDAISY.Addins.Word2003 {
    using System;
    using System.IO;
    using System.Collections;
    using Extensibility;
    using System.Xml;
    using System.Reflection;
    using System.IO.Packaging;
    using System.Windows.Forms;
    using System.Collections.Generic;
    using Microsoft.Office.Core;
    using System.Runtime.InteropServices;
    using MSword = Microsoft.Office.Interop.Word;
    using Daisy.SaveAsDAISY.Conversion;
    using Daisy.SaveAsDAISY;
    using System.Drawing;
    using System.Drawing.Imaging;

    using Microsoft.Win32;
    using System.Text;
    using Word = Microsoft.Office.Interop.Word.InlineShape;
    using IConnectDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
    using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
    using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
    using ConnectIEnumETC = System.Runtime.InteropServices.ComTypes.IEnumFORMATETC;
    using COMException = System.Runtime.InteropServices.COMException;
    using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;
    using Daisy.SaveAsDAISY.Forms;

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the DaisyTranslatorWord2003AddinSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("32B42A31-8592-459F-B6AE-AC70D254CD14"), ProgId("DaisyTranslatorWord2003Addin.Connect")]
	public class Connect : Object, Extensibility.IDTExtensibility2
	{

		const string coreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
		const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
		const string coreNamespace = "http://purl.org/dc/elements/1.1/";
		const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
		const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		String docFile, templateName, path_For_Xml, path_For_Pipeline;
		XmlDocument docTemp;
		private int Word12SaveFormat = -1;
		PackageRelationship pacKRelationship = null;
		private const string Word12Class = "Word12";
		CommandBar _CmdToolBar;
		CommandBarButton _BtnMarkAbbr, _BtnShowManualInWord, _BtnShowManualInDtBook, _BtnManAbbr, _BtnMarkAcr, _BtnManAcr, _BtnHelp, _BtnDaisyStyles, _BtnAbout, _BtnValidate, _BtnAddFootNote, _BtnSettings, _BtnLngSettings;
		CommandBarControl _BtnMasterSub, _BtnSaveDiasy, tool_Single, tool_Multiple, tool_singleXml, tool_SingleScr, tool_MulScr, _helpPopup;
		int _TagCount = 1;
		string _TagString = "AcessibilityTag";
		XmlDocument mergeXmlDoc, currentDocXml, validation_xml;
		ArrayList mergeDoclanguage, openSubdocs;
		frmValidate2003 validateInput;
		public ArrayList docValidation = new ArrayList();
		Pipeline pipe;
		ToolStripMenuItem PipelineMenuItem = null;
		ArrayList objectShapes;
		ArrayList inlineShapes;
		ArrayList imageId;
		ArrayList inlineId;
		ArrayList listmathML;
		private ArrayList footntRefrns;
		string masterSubFlag = "";
		ArrayList langMergeDoc, notTranslatedDoc;
		PackageRelationship packRelationship = null, relationship = null;
		Hashtable multipleMathMl = new Hashtable();
		Hashtable multipleOwnMathMl = new Hashtable();

		/// <summary>
		/// Function to convert word .doc to .docx Format
		/// </summary>
		private void FindWord12SaveFormat()
		{
			try
			{
				MSword.FileConverters converters = this.applicationObject.FileConverters;
				foreach (MSword.FileConverter converter in converters)
				{
					if (converter.ClassName == Word12Class)
					{
						// Converter found
						Word12SaveFormat = converter.SaveFormat;
						break;
					}
				}
			}
			catch
			{
				// No user disturbance...
			}
		}


		/// <summary>
		///		Implements the constructor for the Add-in object.
		///		Place your initialization code within this method.
		/// </summary>
		public Connect()
		{
			this.addinLib = new Daisy.SaveAsDAISY.AddinResources();
		}

		/// <summary>
		///      Implements the OnConnection method of the IDTExtensibility2 interface.
		///      Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param term='application'>
		///      Root object of the host application.
		/// </param>
		/// <param term='connectMode'>
		///      Describes how the Add-in is being loaded.
		/// </param>
		/// <param term='addInInst'>
		///      Object representing this Add-in.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
			this.applicationObject = (MSword.Application)application;
			FindWord12SaveFormat();
			if (connectMode != Extensibility.ext_ConnectMode.ext_cm_Startup)
			{
				OnStartupComplete(ref custom);
			}
		}

		/// <summary>
		///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
		///     Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param term='disconnectMode'>
		///      Describes how the Add-in is being unloaded.
		/// </param>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			if (disconnectMode != Extensibility.ext_DisconnectMode.ext_dm_HostShutdown)
			{
				OnBeginShutdown(ref custom);
			}

			this.applicationObject = null;
			this.addinLib = null;

		}

		/// <summary>
		///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		///      Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		/// <summary>
		///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		///      Receives notification that the host application has completed loading.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref System.Array custom)
		{
			_ResetToolbar();

			CommandBar commandBar = applicationObject.CommandBars["File"];
			// Add import button
			path_For_Pipeline = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms";
			if (Directory.Exists(path_For_Pipeline))
			{
				try
				{
					importButton = (CommandBarControl)commandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 6, true);
				}
				catch (Exception)
				{
					//otherwise, create a new one
					importButton = (CommandBarControl)commandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 6, true);
				}

				importButton.Caption = this.addinLib.GetString("DaisyImportLabel");
				importButton.Tag = this.addinLib.GetString("DaisySaveASLabel");
				importButton.TooltipText = this.addinLib.GetString("DaisySaveASLabel");
				importButton.Visible = true;
				CommandBarPopup popupMenu = (CommandBarPopup)importButton;
				CommandBar barImport = popupMenu.CommandBar;
				single = barImport.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 1, true);
				//single.Caption = this.addinLib.GetString("DaisySingle");
				single.Caption = "Daisy";
				single.Tag = "_postprocess";
				single.Visible = true;
				single.TooltipText = "DAISY XML only (dtbook dtd)";
				CommandBarButton importButtonSingle = (CommandBarButton)single;
				importButtonSingle.Picture = this.addinLib.GetLogo("Singlexml.gif");
				importButtonSingle.Caption = this.addinLib.GetString("DaisySingle");

				importButtonSingle.Click += new _CommandBarButtonEvents_ClickEventHandler(importButtonSingle_Click);

				singleXml = barImport.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 2, true);
				singleXml.Caption = this.addinLib.GetString("DaisyXml");
				singleXml.Visible = true;
				singleXml.TooltipText = "Full text and full audio in MP3";
				CommandBarPopup mnuDtbookSingle = (CommandBarPopup)singleXml;
				CommandBar barDtbook = mnuDtbookSingle.CommandBar;
				pipe = new Pipeline(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms\scripts");
				foreach (KeyValuePair<string, FileInfo> k in pipe.ScriptsInfo)
				{
					PipelineMenuItem = new ToolStripMenuItem();
					PipelineMenuItem.Text = k.Key;
					PipelineMenuItem.AccessibleName = k.Key;
					Singlescript = barDtbook.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 1, true);
					Singlescript.Caption = PipelineMenuItem.Text;
					Singlescript.Visible = true;
					Singlescript.Tag = k.Key;
					CommandBarButton btn_ScrSingle = (CommandBarButton)Singlescript;
					btn_ScrSingle.Picture = this.addinLib.GetLogo("speaker.gif");
					btn_ScrSingle.Click += new _CommandBarButtonEvents_ClickEventHandler(importButtonSingle_Click);
				}


				//Multiple document
				multiple = barImport.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 3, true);
				multiple.Caption = this.addinLib.GetString("DaisyMultiple");
				multiple.Visible = true;
				multiple.Tag = "DaisyMultiple";
				multiple.TooltipText = "DAISY XML only (dtbook dtd)";
				CommandBarButton btn_multipleImport = (CommandBarButton)multiple;
				btn_multipleImport.Picture = this.addinLib.GetLogo("Multiplexml.gif");
				btn_multipleImport.Caption = this.addinLib.GetString("DaisyMultiple");
				btn_multipleImport.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnMasterSub_Click);
				mutipleXml = barImport.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 4, true);
				mutipleXml.Caption = this.addinLib.GetString("DaisyMultipleDtbook");
				mutipleXml.Visible = true;
				mutipleXml.TooltipText = "Full text and full audio in MP3";
				CommandBarPopup mnuDtbookMultiple = (CommandBarPopup)mutipleXml;
				CommandBar barDtbookMul = mnuDtbookMultiple.CommandBar;
				foreach (KeyValuePair<string, FileInfo> k in pipe.ScriptsInfo)
				{
					PipelineMenuItem = new ToolStripMenuItem();
					PipelineMenuItem.Text = k.Key;
					PipelineMenuItem.AccessibleName = k.Key;
					multipleScript = barDtbookMul.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1, true);
					multipleScript.Caption = PipelineMenuItem.Text;
					multipleScript.Visible = true;
					multipleScript.Tag = "DaisyDTBookMultiple";
					CommandBarButton btn_multipleScript = (CommandBarButton)multipleScript;
					btn_multipleScript.Picture = this.addinLib.GetLogo("subfolder.gif");
					btn_multipleScript.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnMasterSub_Click);
				}
			}
			else
			{
				try
				{
					// if item already exists, use it (should never happen)
					import = (CommandBarButton)commandBar.Controls[this.addinLib.GetString("DaisyImportLabel")];
				}
				catch (Exception)
				{
					// otherwise, create a new one
					import = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, 6, true);
				}
				import.Caption = this.addinLib.GetString("DaisyImportLabel");
				import.Tag = this.addinLib.GetString("DaisySaveASLabel");
				import.TooltipText = this.addinLib.GetString("DaisySaveASLabel");
				// set action
				import.OnAction = "!<DaisyTranslatorWord2003Addin.Connect>";
				import.Visible = true;
				import.Picture = this.addinLib.GetLogo("speaker.gif");
				import.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(importButtonSingle_Click);
			}

			// set item's label           
			applicationObject.NormalTemplate.Saved = true;
			_AddCommandBar();
			_DelegateEvents();
		}


		/// <summary>
		/// Function to set delegates and events
		/// </summary>
		private void _DelegateEvents()
		{
			_BtnMarkAbbr.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnMarkAbbr_Click);
			_BtnManAbbr.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnManAbbr_Click);
			_BtnMarkAcr.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnMarkAcr_Click);
			_BtnManAcr.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnManAcr_Click);
			_BtnDaisyStyles.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnDaisyStyles_Click);
			//_BtnHelp.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnHelp_Click);
			_BtnAbout.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnAbout_Click);
			_BtnValidate.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnValidate_Click);
			_BtnAddFootNote.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnAddFootNote_Click);
			_BtnSettings.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnSettings_Click);
			_BtnLngSettings.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnLngSettings_Click);
			_BtnShowManualInWord.Click += new _CommandBarButtonEvents_ClickEventHandler(ShowWordManual);
			_BtnShowManualInDtBook.Click += new _CommandBarButtonEvents_ClickEventHandler(ShowDtBookManual);
			this.applicationObject.DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(applicationObject_DocumentBeforeSave);
			this.applicationObject.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(applicationObject_DocumentOpen);
			this.applicationObject.DocumentChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(applicationObject_DocumentChange);
		}

		void _BtnAddFootNote_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			AddFotenote();
		}

		void _BtnSettings_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			ConverterSettingsForm daisyfrm = new ConverterSettingsForm();
			daisyfrm.ShowDialog();
		}


		void _BtnLngSettings_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			GetlanguageSettings();
		}

		/// <summary>
		/// Core function to Translate the bunch of Documents
		/// </summary>
		/// <param name="control"></param>
		void _BtnMasterSub_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			Application.DoEvents();
			try {
				GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
				IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
				FileInfo pipelineScript = this.pipe?.ScriptsInfo[Ctrl.Tag];

				ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript.FullName);
				WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
				GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
				// Note : the current form for multiple also include conversion settings update
				List<string> documentsPathes = converter.requestUserDocumentsList();
				if (documentsPathes != null && documentsPathes.Count > 0) {
					List<DocumentParameters> documents = new List<DocumentParameters>();

					foreach (string inputPath in documentsPathes) {
						DocumentParameters subDoc = null;
						try {
							subDoc = converter.preprocessDocument(inputPath);
						} catch (Exception e) {
							string errors = "Convertion aborted due to the following errors found while preprocessing " + inputPath + ":\r\n" + e.Message;
							eventsHandler.onPreprocessingError(inputPath, errors);
						}
						if (subDoc != null) {
							documents.Add(subDoc);
						} else {
							// abort documents conversion
							documents.Clear();
							break;
						}
					}
					if (documents.Count > 0) converter.convert(documents);
				}

				applicationObject.ActiveDocument.Save();
			} catch (Exception e) {
				MessageBox.Show(e.Message);
			}
		}

		void applicationObject_DocumentChange()
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			templateName = (doc.get_AttachedTemplate() as MSword.Template).Name;
			CheckforAttchments();
			//Enabling Validate tab
			_BtnValidate.Enabled = true;
		}


		/// <summary>
		///Core Function to check the Daisy Styles in the current Document
		/// </summary>
		/// <param name="Doc">Current Document</param>
		void applicationObject_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			templateName = (doc.get_AttachedTemplate() as MSword.Template).Name;
			CheckforAttchments();
		}

		/// <summary>
		/// Function to check whether the Daisy Styles are new or not
		/// </summary>
		public void CheckforAttchments()
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			Dictionary<int, string> objArray = new Dictionary<int, string>();
			for (int iCountStyles = 1; iCountStyles <= doc.Styles.Count; iCountStyles++)
			{
				object ActualVal = iCountStyles;
				string strValue = doc.Styles.get_Item(ref ActualVal).NameLocal.ToString();
				if (strValue.EndsWith("(DAISY)"))
				{
					objArray.Add(iCountStyles, strValue);
				}
			}

			if (objArray.Count == AddInHelper.DaisyStylesCount)
			{
				_BtnDaisyStyles.Enabled = false;
			}
			else
			{
				_BtnDaisyStyles.Enabled = true;
			}
		}

		/// <summary>
		/// Function to remove Unwanted bookmarks befreo saving the document
		/// </summary>
		/// <param name="Doc"></param>
		/// <param name="SaveAsUI"></param>
		/// <param name="Cancel"></param>
		void applicationObject_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
		{
			DocumentProperties xmlparts = (DocumentProperties)Doc.CustomDocumentProperties;

			foreach (object item in Doc.Bookmarks)
			{
				bool exists = false;
				String bkmrkName = ((MSword.Bookmark)item).Name;

				if (bkmrkName.StartsWith("Abbreviations") || bkmrkName.StartsWith("Acronyms"))
				{
					foreach (DocumentProperty prop in xmlparts)
					{
						if (prop.Name == bkmrkName)
						{
							exists = true;
							String valuebook = prop.Value.ToString();
							String val = valuebook.Substring(0, valuebook.IndexOf("$#$"));
							if (val != ((MSword.Bookmark)item).Range.Text.Trim())
							{
								((MSword.Bookmark)item).Delete();
								xmlparts[prop.Name].Delete();
							}
						}
					}
					if (!exists)
					{
						((MSword.Bookmark)item).Delete();
					}
				}
			}

			foreach (DocumentProperty prop in xmlparts)
			{
				if (prop.Name.StartsWith("Abbreviations") || prop.Name.StartsWith("Acronyms"))
				{
					if (!Doc.Bookmarks.Exists(prop.Name))
					{
						xmlparts[prop.Name].Delete();
					}
				}
			}
		}




		/// <summary>
		/// Function to add Command Bar to the Word Document
		/// </summary>
		private void _AddCommandBar()
		{
			try
			{
				// IF THE COMMAND BAR IS NOT FOUND CREATE A NEW COMMAND BAR 
				if (_CmdToolBar == null)
				{
					_CmdToolBar = (CommandBar)this.applicationObject.CommandBars.Add("Accessibility", MsoBarPosition.msoBarTop, System.Reflection.Missing.Value, (object)true);
					_CmdToolBar.Position = MsoBarPosition.msoBarTop;
				}

				// ADDING THE BUTTONS TO THE COMMAND BAR
				if (_CmdToolBar != null)
				{
					if (Directory.Exists(path_For_Pipeline))
					{
						_BtnSaveDiasy = (CommandBarControl)_CmdToolBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", Type.Missing, true);
						_BtnSaveDiasy.Caption = "SaveAsDAISY";
						_BtnSaveDiasy.Visible = true;
						CommandBarPopup popupMenu = (CommandBarPopup)_BtnSaveDiasy;
						CommandBar barImport = popupMenu.CommandBar;
						tool_Single = barImport.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 1, true);
						tool_Single.Caption = this.addinLib.GetString("DaisySingle");
						tool_Single.Tag = "_postprocess";
						tool_Single.Visible = true;
						tool_Single.TooltipText = "Translate the document to DAISY xml.";
						CommandBarButton importButtonSingle = (CommandBarButton)tool_Single;
						importButtonSingle.Picture = this.addinLib.GetLogo("Singlexml.gif");
						importButtonSingle.Caption = this.addinLib.GetString("DaisySingle");

						tool_singleXml = barImport.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 2, true);
						tool_singleXml.Caption = this.addinLib.GetString("DaisyXml");
						tool_singleXml.Visible = true;

						tool_singleXml.TooltipText = "Translate the document to DAISY DTBook file set.";
						CommandBarPopup mnuDtbookSingle = (CommandBarPopup)tool_singleXml;
						CommandBar barDtbook = mnuDtbookSingle.CommandBar;

						foreach (KeyValuePair<string, FileInfo> k in pipe.ScriptsInfo) if (!k.Key.Equals("_postprocess"))  {
							PipelineMenuItem.Text = k.Key;
							PipelineMenuItem.AccessibleName = k.Key;
							tool_SingleScr = barDtbook.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 1, true);
							tool_SingleScr.Caption = PipelineMenuItem.Text;
							tool_SingleScr.Visible = true;
							tool_SingleScr.Tag = k.Key;
							CommandBarButton btn_ScrSingle = (CommandBarButton)tool_SingleScr;
							btn_ScrSingle.Picture = this.addinLib.GetLogo("speaker.gif");
						}


						//Multiple documents 

						_BtnMasterSub = barImport.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 3, true);
						_BtnMasterSub.Caption = this.addinLib.GetString("DaisyMultiple");
						_BtnMasterSub.Visible = true;
						_BtnMasterSub.Tag = "DaisyMultiple";
						_BtnMasterSub.TooltipText = "Translate multiple documents to DAISY xml.";
						CommandBarButton btn_multipleImport = (CommandBarButton)_BtnMasterSub;
						btn_multipleImport.Picture = this.addinLib.GetLogo("Multiplexml.gif");
						btn_multipleImport.Caption = this.addinLib.GetString("DaisyMultiple");

						tool_Multiple = barImport.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", 4, true);
						tool_Multiple.Caption = this.addinLib.GetString("DaisyMultipleDtbook");
						tool_Multiple.Visible = true;
						tool_Multiple.TooltipText = "Translate multiple documents to DAISY DTBook file set.";
						CommandBarPopup mnuDtbookMultiple = (CommandBarPopup)tool_Multiple;
						CommandBar barDtbookMul = mnuDtbookMultiple.CommandBar;

						foreach (KeyValuePair<string, FileInfo> k in pipe.ScriptsInfo) if (!k.Key.Equals("_postprocess")) {
							PipelineMenuItem.Text = k.Key;
							PipelineMenuItem.AccessibleName = k.Key;
							tool_MulScr = barDtbookMul.Controls.Add(MsoControlType.msoControlButton, Type.Missing, "", 1, true);
							tool_MulScr.Caption = PipelineMenuItem.Text;
							tool_MulScr.Visible = true;
							tool_MulScr.Tag = k.Key;
							CommandBarButton btn_multipleScript = (CommandBarButton)tool_MulScr;
							btn_multipleScript.Picture = this.addinLib.GetLogo("subfolder.gif");
						}

					}
					else
					{
						btnImport = _AddCommandBarButton(_CmdToolBar, "Save as DAIS&Y", true, 230, this.addinLib.GetLogo("speaker.gif"));
						btnImport.Tag = "Button1";
						btnImport.Click += new _CommandBarButtonEvents_ClickEventHandler(importButtonSingle_Click);
						btnMasterSub = _AddCommandBarButton(_CmdToolBar, "Save as DAISY - &Multiple Documents", true, 232, this.addinLib.GetLogo("subfolder.gif"));
						btnMasterSub.Tag = "Button2";
						btnMasterSub.Click += new _CommandBarButtonEvents_ClickEventHandler(_BtnMasterSub_Click);
					}
					//Other toolbar buttons
					_BtnMarkAbbr = _AddCommandBarButton(_CmdToolBar, "Mark as A&bbreviation", true, 231, this.addinLib.GetLogo("abbr1.gif"));
					_BtnManAbbr = _AddCommandBarButton(_CmdToolBar, "Ma&nage Abbreviation", true, 232, this.addinLib.GetLogo("abbr2.gif"));
					_BtnMarkAcr = _AddCommandBarButton(_CmdToolBar, "Mark as A&cronym", true, 231, this.addinLib.GetLogo("acr1.gif"));
					_BtnManAcr = _AddCommandBarButton(_CmdToolBar, "Mana&ge Acronym", true, 232, this.addinLib.GetLogo("acr2.gif"));
					_BtnValidate = _AddCommandBarButton(_CmdToolBar, "&Validate", true, 232, this.addinLib.GetLogo("validate.gif"));
					_BtnDaisyStyles = _AddCommandBarButton(_CmdToolBar, "&Import DAISY Styles", true, 234, this.addinLib.GetLogo("import.gif"));
					_BtnAddFootNote = _AddCommandBarButton(_CmdToolBar, "Add &FootNote", true, 232, this.addinLib.GetLogo("Footnotes.gif"));
					_BtnLngSettings = _AddCommandBarButton(_CmdToolBar, "Language", true, 232, this.addinLib.GetLogo("language.gif"));
					_BtnSettings = _AddCommandBarButton(_CmdToolBar, "Settings", true, 232, this.addinLib.GetLogo("gear.gif"));
					_BtnAbout = _AddCommandBarButton(_CmdToolBar, "Version &Details", true, 235, this.addinLib.GetLogo("version.gif"));
					//_BtnHelp = _AddCommandBarButton(_CmdToolBar, "&Help", true, 232, this.addinLib.GetLogo("help.gif"));

					//HELP Popup
                    _helpPopup = (CommandBarControl)_CmdToolBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, "", _TagCount + 1, true);
                    _helpPopup.Caption = "&Help";
                    _helpPopup.Tag = _TagString + _TagCount.ToString(); ;
                    _helpPopup.TooltipText = "Help";
                    _helpPopup.Visible = true;
                    CommandBarPopup helpPopup = (CommandBarPopup)_helpPopup;
                    CommandBar barHelp = helpPopup.CommandBar;
                    _BtnShowManualInWord = _AddCommandBarButton(barHelp, "DAISY Manual in Word DOCX Format ", true, 232, this.addinLib.GetLogo("help.gif"));
                    _BtnShowManualInDtBook = _AddCommandBarButton(barHelp, "DAISY Manual in DTBook Format ", true, 232, this.addinLib.GetLogo("help.gif"));
				}
				_CmdToolBar.Enabled = true;
				_CmdToolBar.Visible = true;

			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message);
			}
		}
		/// <summary>
		/// Function to show the current version of the SaveAsDAISY Addin
		/// </summary>
		/// <param name="control"></param>
		private void _BtnAbout_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			About abtForm = new About();
			abtForm.ShowDialog();
		}

		/// <summary>
		/// Function to show Help File
		/// </summary>
		/// <param name="control"></param>
		private void _BtnHelp_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm");
		}

		public void ShowWordManual(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Daisy_Translator_Instruction_Manual_01_March_2011_v2_5.docx");
		}

		public void ShowDtBookManual(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Save-as-instruction-manual");
		}

		/// <summary>
		/// Function to update the Daisy Styles
		/// </summary>
		/// <param name="control"></param>
		private void _BtnDaisyStyles_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			object copyTempName = null;
			MSword.Document docTemplate = new Microsoft.Office.Interop.Word.Document();
			MSword.Document docActive = this.applicationObject.ActiveDocument;
			object missing = Type.Missing;
			try
			{
				string templateName = (docActive.get_AttachedTemplate() as MSword.Template).Name;
				string templatePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "template2003.dot";
				object newTemp = false;
				object documentType = Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument;
				object visble = false;

				copyTempName = Path.GetTempFileName() + ".dotx";

				if (!File.Exists(copyTempName.ToString()))
				{
					File.Copy(templatePath, copyTempName.ToString());
				}

				object tempPath = (object)templatePath;
				docTemplate = applicationObject.Documents.Add(ref tempPath, ref newTemp, ref documentType, ref visble);


				foreach (MSword.Style styleObj in docTemplate.Styles)
				{
					if (styleObj.NameLocal.EndsWith("(DAISY)"))
					{
						this.applicationObject.OrganizerCopy(copyTempName.ToString(), docActive.FullName, styleObj.NameLocal, MSword.WdOrganizerObject.wdOrganizerObjectStyles);
					}
				}
				object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
				docTemplate.Close(ref saveChanges, ref missing, ref missing);

				this.applicationObject.NormalTemplate.Save();

				if (File.Exists(copyTempName.ToString()))
				{
					File.Delete(copyTempName.ToString());
				}

				_BtnDaisyStyles.Enabled = false;
			}
			catch (Exception ex)
			{
				string stre = ex.Message;
				MessageBox.Show(ex.Message.ToString(), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
				docTemplate.Close(ref saveChanges, ref missing, ref missing);
			}
		}

		/// <summary>
		/// Function to crete Button under Accessibilty Tool Bar 
		/// </summary>
		/// <param name="_pMenuItem">Menu for the Button</param>
		/// <param name="_pCaption"> caption</param>
		/// <param name="_pBeginGroup"></param>
		/// <param name="_pFaceId">Face Id for the button</param>
		/// <param name="image">Image for the button</param>
		/// <returns></returns>
		public CommandBarButton _AddCommandBarButton(CommandBar _pMenuItem, string _pCaption, bool _pBeginGroup, int _pFaceId, stdole.IPictureDisp image)
		{
			try
			{
				CommandBarButton _lTempButton = null;
				_lTempButton = (CommandBarButton)_pMenuItem.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
				_lTempButton.Caption = _pCaption;
				_lTempButton.BeginGroup = _pBeginGroup;
				_lTempButton.Parameter = _pCaption;
				_lTempButton.Style = MsoButtonStyle.msoButtonIcon;
				_TagCount++;
				_lTempButton.Tag = _TagString + _TagCount.ToString();
				_lTempButton.Picture = image;
				//if (_pFaceId != 0)
				//    _lTempButton.FaceId = _pFaceId;
				return _lTempButton;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// Function to crete Button under Accessibilty Tool Bar 
		/// </summary>
		/// <param name="_pMenuItem">Menu</param>
		/// <param name="_pCaption">Caption</param>
		/// <param name="_pBeginGroup"></param>
		/// <param name="_pFaceId">Face Id</param>
		/// <returns></returns>
		public CommandBarButton _AddCommandBarButton(CommandBar _pMenuItem, string _pCaption, bool _pBeginGroup, int _pFaceId)
		{
			try
			{
				CommandBarButton _lTempButton = null;
				_lTempButton = (CommandBarButton)_pMenuItem.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
				_lTempButton.Caption = _pCaption;
				_lTempButton.BeginGroup = _pBeginGroup;
				_lTempButton.Parameter = _pCaption;
				_lTempButton.Style = MsoButtonStyle.msoButtonIcon;
				_TagCount++;
				_lTempButton.Tag = _TagString + _TagCount.ToString();
				if (_pFaceId != 0)
					_lTempButton.FaceId = _pFaceId;
				return _lTempButton;
			}
			catch (Exception ex)
			{
				throw ex;
			}

		}

		/*Function which Translates docx file to Daisy format */
		private void importButtonSingle_Click(CommandBarButton Ctrl, ref bool CancelDefault)
		{
			objectShapes = new ArrayList();
			multipleMathMl = new Hashtable();
			imageId = new ArrayList();
			inlineShapes = new ArrayList();
			inlineId = new ArrayList();
			listmathML = new ArrayList();
			int fileIndex;
			MSword.Document doc = this.applicationObject.ActiveDocument;
			if (!doc.Saved || doc.FullName.LastIndexOf('.') < 0) {
				System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforeExport"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

			} else if (doc.Saved) {
				fileIndex = doc.FullName.LastIndexOf('.');
				String substr = doc.FullName.Substring(fileIndex);

				if (substr.ToLower() != ".docx") {
					System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
				} else {
					try {
						GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
						IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
						FileInfo pipelineScript = this.pipe?.ScriptsInfo[Ctrl.Tag];

						ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript.FullName);
						WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
						GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
						DocumentParameters currentDocument = converter.preprocessDocument(this.applicationObject.ActiveDocument.FullName);
						if (converter.requestUserParameters(currentDocument) == ConversionStatus.ReadyForConversion) {
							ConversionResult result = converter.convert(currentDocument);
							/*if (!(result.Canceled || result.Succeeded)) {
								MessageBox.Show(result., "Conversion aborted");
							}*/
						}

						//applicationObject.ActiveDocument.Save();
					} catch (Exception e) {
						MessageBox.Show(e.Message);
					}
				}
			}

		}

		/// <summary>
		///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		///      Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param term='custom'>
		///      Array of parameters that are host application specific.
		/// </param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref System.Array custom)
		{
			try
			{
				if (Directory.Exists(path_For_Pipeline))
				{
					CommandBarControl button = (CommandBarControl)applicationObject.CommandBars.FindControl(MsoControlType.msoControlPopup, Type.Missing, this.addinLib.GetString("DaisySaveASLabel"), Type.Missing);
					button.Delete(Type.Missing);
				}
				else
				{
					CommandBarButton button = (CommandBarButton)applicationObject.CommandBars.FindControl(MsoControlType.msoControlButton, Type.Missing, this.addinLib.GetString("DaisySaveASLabel"), Type.Missing);
					button.Delete(Type.Missing);
				}
				_ResetToolbar();
				applicationObject.NormalTemplate.Save();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		/*Core Function which Translates Docx file*/
		/*public void OoxToDaisyUI(string inputFile, string tempInputFile, String tempInput, CommandBarButton control, String masterSubFlag)
		{

			if (Directory.Exists(path_For_Pipeline))
			{
				try
				{
					ConversionParameters parameters = new ConversionParameters();
					parameters.InputFile = inputFile;
					parameters.TempInputFile = tempInput;
					parameters.Version = this.applicationObject.Version;
					parameters.ControlName = control.Tag;
					parameters.ObjectShapes = objectShapes;
					parameters.ListMathMl = multipleMathMl;
					parameters.ImageIds = imageId;
					parameters.InlineShapes = inlineShapes;
					parameters.InlineIds = inlineId;
					parameters.ParseSubDocuments = masterSubFlag;
					if (control.Tag == "DaisySingle" || control.Tag == "Button1")
					{
						this.addinLib.StartSingleWordConversion(parameters);
					}
					else
					{
						parameters.ScriptPath = pipe.ScriptsInfo[control.Caption].FullName;
						//parameters.Directory = string.Empty;
						this.addinLib.StartSingleWordConversion(parameters);
					}
				}
				catch (Exception e)
				{
					MessageBox.Show(e.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			else
			{
				ConversionParameters parameters = new ConversionParameters();
				parameters.InputFile = inputFile;
				parameters.TempInputFile = tempInput;
				parameters.Version = this.applicationObject.Version;
				parameters.ControlName = string.Empty;
				parameters.ObjectShapes = objectShapes;
				parameters.ListMathMl = multipleMathMl;
				parameters.ImageIds = imageId;
				parameters.InlineShapes = inlineShapes;
				parameters.InlineIds = inlineId;
				parameters.ParseSubDocuments = masterSubFlag;

				this.addinLib.StartSingleWordConversion(parameters);
			}
		}*/

		private MSword.Application applicationObject;
		private CommandBarControl importButton, single, multiple, singleXml, Singlescript, mutipleXml, multipleScript;
		private CommandBarButton import, btnImport, btnMasterSub;
		private AddinResources addinLib;


		#region Document Properties

		/// <summary>
		/// Function to get the Title of the Current Document
		/// </summary>
		/// <returns>Title</returns>
		public String DocPropTitle()
		{
			int titleFlag = 0;
			String styleVal = "";
			string msgConcat = "";
			Package pack;
			pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
			String strTitle = pack.PackageProperties.Title;
			pack.Close();
			if (strTitle != "")
				return strTitle;
			else
			{
				pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
				foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
				{
					pacKRelationship = searchRelation;
					break;
				}
				Uri partUri = PackUriHelper.ResolvePartUri(pacKRelationship.SourceUri, pacKRelationship.TargetUri);
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
									msgConcat = msgConcat + " " + getText.InnerText;
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

				strTitle = msgConcat;

			}
			return strTitle;
		}

		/// <summary>
		/// Function to get the Creator of the Current Document
		/// </summary>
		/// <returns>Creator</returns>
		public String DocPropCreator()
		{
			Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
			String strCreator = pack.PackageProperties.Creator;
			pack.Close();
			if (strCreator != "")
				return strCreator;
			else
				return "";
		}

		/// <summary>
		///  Function to get the Publisher of the Current Document
		/// </summary>
		/// <returns>Publisher</returns>
		public String DocPropPublish()
		{
			Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
			foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(appRelationshipType))
			{
				pacKRelationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(pacKRelationship.SourceUri, pacKRelationship.TargetUri);
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

		#endregion

		/// <summary>
		/// Resets the toolbar
		/// </summary>
		public void _ResetToolbar()
		{
			try
			{
				CommandBar _pCmdToolBar = _GetToolbar();

				if (_pCmdToolBar != null)
				{
					_pCmdToolBar.Delete();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		/// <summary>
		/// Function to retuen Tool Bar
		/// </summary>
		/// <returns></returns>
		public CommandBar _GetToolbar()
		{
			CommandBar _CmdToolBarTemp = null;
			foreach (CommandBar ComBar in this.applicationObject.CommandBars)
			{
				if (ComBar.Name == "Accessibility")
				{
					_CmdToolBarTemp = ComBar;
					break;
				}
			}
			return _CmdToolBarTemp;
		}

		#region Abbreviation and Acronyms

		/// <summary>
		/// Core function to Mark Abbreviations
		/// </summary>
		/// <param name="control"></param>
		private void _BtnMarkAbbr_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			try
			{
				MSword.Document doc = this.applicationObject.ActiveDocument;
				if (this.applicationObject.Selection.Start == this.applicationObject.Selection.End || this.applicationObject.Selection.Text.Equals("\r"))
				{
					MessageBox.Show(addinLib.GetString("AbbreviationText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (validateSpace(this.applicationObject.Selection.Text.Trim()))
				{
					MessageBox.Show(addinLib.GetString("AbbreviationproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateWord(this.applicationObject.Selection))
				{
					MessageBox.Show(addinLib.GetString("AbbreviationproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateBookMark(this.applicationObject.Selection) == "Abbrtrue")
				{
					MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateBookMark(this.applicationObject.Selection) == "Acrtrue")
				{
					MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					Mark2003 mrkForm = new Mark2003(doc, true);
					mrkForm.ShowDialog();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		/// <summary>
		/// Core Function for Mark Acronyms
		/// </summary>
		/// <param name="control"></param>
		private void _BtnMarkAcr_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			try
			{
				MSword.Document doc = this.applicationObject.ActiveDocument;
				if (this.applicationObject.Selection.Start == this.applicationObject.Selection.End || this.applicationObject.Selection.Text.Equals("\r"))
				{
					MessageBox.Show(addinLib.GetString("AcronymText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (validateSpace(this.applicationObject.Selection.Text.Trim()))
				{
					MessageBox.Show(addinLib.GetString("AcronymproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateWord(this.applicationObject.Selection))
				{
					MessageBox.Show(addinLib.GetString("AcronymproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateBookMark(this.applicationObject.Selection) == "Abbrtrue")
				{
					MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else if (ValidateBookMark(this.applicationObject.Selection) == "Acrtrue")
				{
					MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					Mark2003 mrkForm = new Mark2003(doc, false);
					mrkForm.ShowDialog();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		/// <summary>
		/// Function to check whether the selected Bookmark is Abbreviation\Acronym
		/// </summary>
		/// <param name="select">Range of text</param>
		/// <returns></returns>
		public String ValidateBookMark(MSword.Selection select)
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			String value = "Nobookmarks";

			if (this.applicationObject.Selection.Bookmarks.Count > 0)
			{
				foreach (object item in doc.Bookmarks)
				{
					if (((MSword.Bookmark)item).Range.Text.Trim() == this.applicationObject.Selection.Text.Trim())
					{
						if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
							value = "Abbrtrue";
						if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
							value = "Acrtrue";
					}
				}
			}

			return value;
		}

		/// <summary>
		/// Functio to check the selected Text is valid for Abbreviation\Acronym or not
		/// </summary>
		/// <param name="text"></param>
		/// <returns></returns>
		public bool validateSpace(String text)
		{
			if (text.Contains(" "))
				return true;
			else if (text == "")
				return true;
			else
				return false;
		}

		/// <summary>
		/// Function to check the selected Text is valid for Abbreviation\Acronym or not
		/// </summary>
		/// <param name="select">Range of text</param>
		/// <returns></returns>
		public bool ValidateWord(MSword.Selection select)
		{
			bool flag = false;
			MSword.Document doc = this.applicationObject.ActiveDocument;

			if (select.Words.Count == 1)
			{
				if (this.applicationObject.Selection.Words.Last.Text.Trim() != this.applicationObject.Selection.Text.Trim())
					flag = true;
			}
			return flag;
		}

		/// <summary>
		/// Core function to Manage Abbreviations
		/// </summary>
		/// <param name="control"></param>
		private void _BtnManAbbr_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			try
			{
				MSword.Document doc = this.applicationObject.ActiveDocument;
				if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
				{
					Abbreviation2003 abbrForm = new Abbreviation2003(doc, true);
					abbrForm.ShowDialog();
				}
				else
				{
					MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		/// <summary>
		/// Core function to Manage Acronyms
		/// </summary>
		/// <param name="control"></param>
		private void _BtnManAcr_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			try
			{
				MSword.Document doc = this.applicationObject.ActiveDocument;
				if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
				{
					Abbreviation2003 AcrForm = new Abbreviation2003(doc, false);
					AcrForm.ShowDialog();
				}
				else
				{
					MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		#endregion


		#region Multiple OOXML functions

		/// <summary>
		/// Function to Check selected Documents are Master/Sub dcouemnts or simple Documents 
		/// </summary>
		/// <param name="listSubDocs">Seleted Documents</param>
		/// <param name="value">bool value</param>
		/// <returns>document type</returns>
		public String CheckingSubDocs(ArrayList listSubDocs)
		{
			String resultSubDoc = "simple";
			for (int i = 0; i < listSubDocs.Count; i++)
			{
				string[] splitName = listSubDocs[i].ToString().Split('|');
				Package pack;
				pack = Package.Open(splitName[0].ToString(), FileMode.Open, FileAccess.ReadWrite);

				foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
				{
					packRelationship = searchRelation;
					break;
				}

				Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
				PackagePart mainPartxml = pack.GetPart(partUri);

				foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
				{
					packRelationship = searchRelation;
					if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
					{
						if (packRelationship.TargetMode.ToString() == "External")
						{
							resultSubDoc = "complex";
						}
					}
				}
				pack.Close();
			}
			return resultSubDoc;
		}

		/// <summary>
		/// Function to Check selected Documents are Master/Sub dcouemnts or simple Documents 
		/// </summary>
		/// <param name="listSubDocs">Seleted Documents</param>
		/// <param name="value">bool value</param>
		/// <returns>document type</returns>
		public String CheckingSubDocs(ArrayList listSubDocs, bool value)
		{
			String resultSubDoc = "simple";
			for (int i = 1; i < listSubDocs.Count; i++)
			{
				string[] splitName = listSubDocs[i].ToString().Split('|');
				Package pack;
				pack = Package.Open(splitName[0].ToString(), FileMode.Open, FileAccess.ReadWrite);

				foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
				{
					packRelationship = searchRelation;
					break;
				}

				Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
				PackagePart mainPartxml = pack.GetPart(partUri);

				foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
				{
					packRelationship = searchRelation;
					if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
					{
						if (packRelationship.TargetMode.ToString() == "External")
						{
							resultSubDoc = "complex";
						}
					}
				}
				pack.Close();
			}
			return resultSubDoc;
		}

		public void MathMLMultiple(ArrayList subList)
		{
			multipleMathMl.Clear();
			for (int i = 0; i < subList.Count; i++)
			{
				string[] splitName = subList[i].ToString().Split('|');
				multipleOwnMathMl = new Hashtable();
				MTGetEquationAddinNew(splitName[0].ToString());
				multipleMathMl.Add("Doc" + i, multipleOwnMathMl);
			}
		}

		/// <summary>
		/// Function to whether the Document is already open or Not
		/// </summary>
		/// <param name="listSubDocs">Selected Documents</param>
		/// <returns>message</returns>
		public String CheckFileOPen(ArrayList listSubDocs)
		{
			String resultSubDoc = "notopen";
			openSubdocs = new ArrayList();
			for (int i = 0; i < listSubDocs.Count; i++)
			{
				string[] splt = listSubDocs[i].ToString().Split('|');
				try
				{
					Package pack;
					pack = Package.Open(splt[0].ToString(), FileMode.Open, FileAccess.ReadWrite);
					pack.Close();
				}
				catch (Exception e)
				{
					resultSubDoc = "open";
					openSubdocs.Add(splt[0].ToString() + "\n");
				}
			}
			return resultSubDoc;
		}

		#endregion

		private void _BtnValidate_Click(CommandBarButton Ctrl, ref Boolean CancelDefault)
		{
			//creating instance of arraylist docValidation
			docValidation = new ArrayList();
			int ind;
			MSword.Document doc = this.applicationObject.ActiveDocument;
			//Path of the active document
			string path = doc.FullName;
			//Checking whether document is saved
			if (!doc.Saved || doc.FullName.LastIndexOf('.') < 0)
			{
				System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforeValidate"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
				_BtnValidate.Enabled = true;
			}
			//if saved
			else if (doc.Saved)
			{
				ind = doc.FullName.LastIndexOf('.');
				String substr = doc.FullName.Substring(ind);

				if (substr.ToLower() != ".docx")
				{
					System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
					_BtnValidate.Enabled = true;
				}
				else
				{

					object newName = Path.GetTempFileName() + Path.GetExtension((string)doc.FullName);
					File.Copy((string)doc.FullName, (string)newName);

					// open the duplicated file
					object addToRecentFiles = false;
					object readOnly = false;
					object isVisible = false;
					object missing = Type.Missing;

					MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
					// generate docx file from the duplicated file (under a temporary file)
					object tmpFileName = this.addinLib.GetTempPath((string)doc.FullName, ".docx");
					object format = Word12SaveFormat;
					newDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

					// close and remove the duplicated file
					object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
					object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
					newDoc.Close(ref saveChanges, ref originalFormat, ref missing);

					try
					{
						File.Delete((string)newName);
					}
					catch (IOException)
					{

					}
					docFile = (string)tmpFileName;
				}
				DeleteBookMark();

				//Getting validation xml from executing assembly
				path_For_Xml = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY";

				validation_xml = new XmlDocument();
				//chcking for validation xml
				if (File.Exists(path_For_Xml + "\\prepopulated_validationVer21.xml"))
				{
					validation_xml.Load(path_For_Xml + "\\prepopulated_validationVer21.xml");
				}
				//Opening package
				Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
				foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
				{
					pacKRelationship = searchRelation;
					break;
				}
				Uri partUri = PackUriHelper.ResolvePartUri(pacKRelationship.SourceUri, pacKRelationship.TargetUri);
				PackagePart mainPartxml = pack.GetPart(partUri);
				currentDocXml = new XmlDocument();
				currentDocXml.Load(mainPartxml.GetStream());
				pack.Close();
				//Progress bar
				using (ProgressValidation prog = new ProgressValidation(currentDocXml, validation_xml))
				{
					//Progress bar dialog
					DialogResult dr = prog.ShowDialog();
					//If work complete
					if (dr == System.Windows.Forms.DialogResult.OK)
					{
						docValidation = prog.GetFileNames;
						//If docValidation array list contains validation error
						if (docValidation.Count != 0)
						{
							//Closing the word document
							object saveChange = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
							object originalFmt = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
							object FileName = (object)path;
							object AddToRecentFiles = false;
							object routeDocument = Type.Missing;
							object Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
							MessageBox.Show("This document contains validation errors", "Validation completed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
							this.applicationObject.ActiveDocument.Close(ref saveChange, ref originalFmt, ref routeDocument);
							//Opening the package of the closed document
							Package packClose = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
							foreach (PackageRelationship searchRelation in packClose.GetRelationshipsByType(wordRelationshipType))
							{
								pacKRelationship = searchRelation;
								break;
							}

							Uri packClosePartUri = PackUriHelper.ResolvePartUri(pacKRelationship.SourceUri, pacKRelationship.TargetUri);
							PackagePart packCloseMainPartxml = packClose.GetPart(packClosePartUri);
							currentDocXml = prog.XmlFileNames;
							using (Stream str = packCloseMainPartxml.GetStream(FileMode.Create))
							{
								using (StreamWriter ts = new StreamWriter(str))
								{
									ts.Write(currentDocXml.OuterXml);
									ts.Flush();
									ts.Close();
								}
							}
							packClose.Flush();
							packClose.Close();
							object formatWord = Word12SaveFormat;
							object visible = true;
							object repair = false;
							//Opening the word document
							this.applicationObject.Documents.Open(ref FileName, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument, ref visible, ref routeDocument, ref routeDocument, ref routeDocument, ref routeDocument);
							//Initializing frmValidate2003 form
							validateInput = new frmValidate2003();
							//Calling method and passing validation xml and active document
							validateInput.SetApp(docValidation, this.applicationObject.ActiveDocument);
							//Showing frmValidate2003 form
							validateInput.Show(new WordWin32Window(this.applicationObject));
							//Event on close of frmValidate2003 form
							validateInput.FormClosed += new FormClosedEventHandler(validateInput_FormClosed);
							//Disabling validation tab
							_BtnValidate.Enabled = false;
						}
						else
						{
							//Message if document is valid
							MessageBox.Show("This document is valid", "Validation completed", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
							//Enabling Validate tab
							_BtnValidate.Enabled = true;

						}
					}
					//Progress bar canceled
					if (dr == System.Windows.Forms.DialogResult.Cancel)
					{
						MessageBox.Show("Validation stopped", "Quit validation", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
						_BtnValidate.Enabled = true;
					}
				}
			}
		}

		/*Function which deletes bookmark */
		public void DeleteBookMark()
		{
			try
			{
				MSword.Document doc = this.applicationObject.ActiveDocument;
				//Looping through the document for each Bookmark
				foreach (object item in doc.Bookmarks)
				{
					//Checking for bookmark Rule
					if (((MSword.Bookmark)item).Name.StartsWith("Rule", StringComparison.CurrentCulture))
					{
						object val = ((MSword.Bookmark)item).Name;
						doc.Bookmarks.get_Item(ref val).Delete();
					}
				}
				//Saving the active word document
				doc.Save();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		/*Function which deletes bookmarks on close of validateform */
		void validateInput_FormClosed(object sender, FormClosedEventArgs e)
		{
			try
			{
				DeleteBookMark();
				_BtnValidate.Enabled = true;
				this.applicationObject.ActiveDocument.Save();
			}
			catch
			{

			}
		}

		//Function to generate random ID
		public long GenerateId()
		{
			byte[] buffer = Guid.NewGuid().ToByteArray();
			return BitConverter.ToInt64(buffer, 0);
		}

		#region ShapesObject

		public void saveasshapes(String masterSubFlag)
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
			foreach (MSword.Shape item in doc.Shapes)
			{
				if (!item.Name.Contains("Text Box"))
				{
					object missing = Type.Missing;
					item.Select(ref missing);
					this.applicationObject.Selection.CopyAsPicture();
					Metafile metaFile = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
					System.Drawing.Image image = metaFile;
					byte[] Ret;
					MemoryStream ms = new MemoryStream();
					image.Save(ms, ImageFormat.Png);
					Ret = ms.ToArray();
					String fileName = doc.Name.Replace(" ", "_");
					string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-Shape" + item.ID.ToString() + ".png";
					FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
					fs.Write(Ret, 0, Ret.Length);
					fs.Flush();
					fs.Dispose();
					Clipboard.Clear();
				}
			}
			if (masterSubFlag != "Yes")
				this.applicationObject.ActiveDocument.Save();
		}

		public void saveasshapes(ArrayList subList, String masterSubFlag)
		{
			object addToRecentFiles = false;
			object readOnly = false;
			object isVisible = false;
			object missing = Type.Missing;
			object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
			object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;

			System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
			for (int i = 0; i < subList.Count; i++)
			{
				object newName = subList[i];
				MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
				foreach (MSword.Shape item in newDoc.Shapes)
				{
					if (!item.Name.Contains("Text Box"))
					{
						item.Select(ref missing);
						this.applicationObject.Selection.CopyAsPicture();
						Metafile metaFile = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
						System.Drawing.Image image = metaFile;
						byte[] Ret;
						MemoryStream ms = new MemoryStream();
						image.Save(ms, ImageFormat.Png);
						Ret = ms.ToArray();
						String fileName = newName.ToString().Replace(" ", "_");
						string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-Shape" + item.ID.ToString() + ".png";
						FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
						fs.Write(Ret, 0, Ret.Length);
						fs.Flush();
						fs.Dispose();
						Clipboard.Clear();
					}
				}
				newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
			}
		}

		public void SaveasImages(String masterSubFlag)
		{
			MSword.Document doc = this.applicationObject.ActiveDocument;
			object missing = Type.Missing;
			System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();

			MSword.Range rng;
			foreach (MSword.Range tmprng in doc.StoryRanges)
			{
				rng = tmprng;
				while (rng != null)
				{
					foreach (MSword.InlineShape item in rng.InlineShapes)
					{
						if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture")))
						{
							object range = item.Range;
							string str = "Shapes_" + GenerateId().ToString();
							item.Range.Bookmarks.Add(str, ref range);
							item.Range.CopyAsPicture();
							Metafile metaFile = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
							System.Drawing.Image image = metaFile;
							byte[] Ret;
							MemoryStream ms = new MemoryStream();
							image.Save(ms, ImageFormat.Png);
							Ret = ms.ToArray();
							String fileName = doc.Name.Replace(" ", "_");
							string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png";
							FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
							fs.Write(Ret, 0, Ret.Length);
							fs.Flush();
							fs.Dispose();
							Clipboard.Clear();
						}
					}
					rng = rng.NextStoryRange;
				}
			}
			if (masterSubFlag != "Yes")
				this.applicationObject.ActiveDocument.Save();
		}

		public void SaveasImages(ArrayList subList, String masterSubFlag)
		{
			System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
			object addToRecentFiles = false;
			object readOnly = false;
			object isVisible = false;
			object missing = Type.Missing;
			object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
			object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
			for (int i = 0; i < subList.Count; i++)
			{
				object newName = subList[i];
				MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
				MSword.Range rng;
				foreach (MSword.Range tmprng in newDoc.StoryRanges)
				{
					rng = tmprng;
					while (rng != null)
					{
						foreach (MSword.InlineShape item in rng.InlineShapes)
						{
							if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture")))
							{
								object range = item.Range;
								string str = "Shapes_" + GenerateId().ToString();
								item.Range.Bookmarks.Add(str, ref range);
								item.Range.CopyAsPicture();
								Metafile metaFile = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
								System.Drawing.Image image = metaFile;
								byte[] Ret;
								MemoryStream ms = new MemoryStream();
								image.Save(ms, ImageFormat.Png);
								Ret = ms.ToArray();
								String fileName = newName.ToString().Replace(" ", "_");
								string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png";
								FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
								fs.Write(Ret, 0, Ret.Length);
								fs.Flush();
								fs.Dispose();
								Clipboard.Clear();
							}
						}
						rng = rng.NextStoryRange;
					}
				}
				newDoc.Save();
				newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
			}

		}
		#endregion

		#region FootNote

		//public void AddFotenote(IRibbonControl control)
		public void AddFotenote()
		{
			try
			{
				MSword.Document docActive = this.applicationObject.ActiveDocument;

				Microsoft.Office.Interop.Word.Range footnotetext = docActive.Application.Selection.Range;

				if (footnotetext.Text == null)
				{
					MessageBox.Show("Please select Footnote text from the document", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					String selectedText = docActive.Application.Selection.Text;
					String refNumber = string.Empty;
					String ftNtText = string.Empty;
					Object refNumberObj = null;

					//Shaby(Fix Start) -  to consider any text provided as index of reference
					//string[] footNote = selectedText.Split('.');
					string[] footNote = selectedText.Split(' ');
					if (footNote.Length <= 1)
					{
						throw new Exception("Selected Text is not a valid Footnote");
					}
					refNumber = footNote[0];
					//Remove dot at the end, if any
					refNumber = refNumber.TrimEnd('.');
					refNumberObj = refNumber;

					ftNtText = string.Empty;
					for (int i = 1; i < footNote.Length; i++)
					{
						ftNtText = ftNtText + footNote[i] + " ";
					}
					Object actualText = ftNtText;
					ftNtText = actualText.ToString();
					//Shaby(Fix End)

					Object startValue = 0;
					//Fix for searching only the upper part of the body
					Object endValue = footnotetext.Start;//Object endValue = docActive.Content.End;


					MSword.Range rngDoc = docActive.Range(ref startValue, ref endValue);
					//Commented below line as a fix for Bug 6633
					//docActive.Application.Selection.ClearFormatting();
					MSword.Find fndValue = rngDoc.Find;
					fndValue.Forward = false;
					fndValue.Text = refNumber.ToLower();

					//Shaby (start) : Search for string at the end of a word or as a whole string
					//Feature not supported in Word2003
					//fndValue.MatchSuffix = true;
					//Shaby (end)

					footntRefrns = new ArrayList();
					ExecuteFind(fndValue);
					int occurance = 0;
					while (fndValue.Found)
					{
						int Strng = rngDoc.Start;
						Object startRng;
						startRng = Strng - 30;
						if (Convert.ToInt32(startRng) <= 0)
						{
							startRng = 0;
						}
						int rdRng = rngDoc.End;
						Object endRng = rdRng + 20;
						if (Convert.ToInt32(endRng) >= Convert.ToInt32(endValue))
						{
							endRng = endValue;
						}
						MSword.Range getRng = docActive.Range(ref startRng, ref endRng);

						if (getRng.Text != null)
						{
							//Shaby : Below check done to avoid selecting the actual footer
							if (!rngDoc.InRange(footnotetext))
							{
								//Shaby : select the apt index, if there are multiple Footnote index in selected body text 
								if (getRng.Text.ToLower().IndexOf(fndValue.Text) != getRng.Text.ToLower().LastIndexOf(fndValue.Text))
								{
									//Trimming the left position if any index exist in the left part
									int leftDistanceToRef = rngDoc.Start - getRng.Start;
									string leftTextToRef = getRng.Text.Substring(0, leftDistanceToRef);
									if (leftTextToRef.ToLower().Contains(fndValue.Text.ToLower()))
									{
										startRng = rngDoc.Start;
									}
									//Trimming the right position if any index exists in the right part
									int rightDistanceFromRef = getRng.End - rngDoc.End;
									int leftDistanceFromRef = rngDoc.End - getRng.Start;
									string rightTextFromRef = getRng.Text.Substring(leftDistanceFromRef, rightDistanceFromRef);
									if (rightTextFromRef.ToLower().Contains(fndValue.Text.ToLower()))
									{
										endRng = rngDoc.End;
									}
								}

								////Shaby2: Trim till the last linebreak(if any), at both the sides
								MSword.Range lineBreakerRng = docActive.Range(ref startRng, ref endRng);
								while ((lineBreakerRng.Text.IndexOf("\r")) != -1)
								{
									int indexOfFooterIndex = lineBreakerRng.Text.ToLower().IndexOf(fndValue.Text);
									if ((lineBreakerRng.Text.IndexOf("\r")) < indexOfFooterIndex)
									{
										startRng = lineBreakerRng.Start + lineBreakerRng.Text.IndexOf("\r") + 1;
									}
									else
									{
										endRng = lineBreakerRng.Start + lineBreakerRng.Text.IndexOf("\r");
									}

									lineBreakerRng = docActive.Range(ref startRng, ref endRng);
								}

								//Shaby : Below check required to filter only SuperScripts
								//if(getRng.Font.Superscript == 9999999)
								occurance = occurance + 1;
								footntRefrns.Add(getRng.Text + "|" + startRng + "|" + endRng + "|" + refNumberObj + "|" + occurance);
							}
						}

						ExecuteFind(fndValue);

					}


					//Shaby2
					if (footntRefrns.Count != 0)
					{
						SuggestedReferences suggestForm = new SuggestedReferences(docActive, footntRefrns, ftNtText);
						suggestForm.ShowDialog();
					}
					else
					{
						MessageBox.Show("No match found in the document to map selected footnote", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception x)
			{
				MessageBox.Show(x.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		public Boolean ExecuteFind(MSword.Find find)
		{
			object missing = System.Type.Missing;
			object text = find.Text;
			object MatchCase = false;
			object MatchSoundsLike = false;
			object MatchAllWordForms = false;
			object MatchWholeWord = false;
			object MatchWildcards = false;
			object Forward = false;
			object format = false;

			return find.Execute(ref text, ref MatchCase, ref MatchWholeWord,
			  ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms,
			  ref Forward, ref missing, ref format, ref missing, ref missing,
			  ref missing, ref missing, ref missing,
			  ref missing);

		}

		#endregion FootNote

		#region MAthML

		private IConnectDataObject mDataObject;

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		private static extern IntPtr GlobalLock(HandleRef handle);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		private static extern bool GlobalUnlock(HandleRef handle);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		private static extern int GlobalSize(HandleRef handle);

		//#endregion Kernel32 function calls

		// #region OLE32 function calls
		[DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

		[DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
		private static extern int OleGetAutoConvert(ref Guid oCurrentCLSID, out Guid pConvertedClsid);

		[DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern bool IsEqualGUID(ref Guid rclsid1, ref Guid rclsid);

		private void MTGetEquationAddin()
		{
			Microsoft.Office.Interop.Word.Document doc = this.applicationObject.ActiveDocument;

			if (doc != null)
			{
				int iNumShapesIterated = 0;
				iNumShapesIterated = IterateShapes();
			}
		}

		private int IterateShapes()
		{
			Int16 showMsg = 0;
			MSword.Range rng;
			String storyName = "";
			int iNumShapesViewed = 0;
			Microsoft.Office.Interop.Word.Document doc = this.applicationObject.ActiveDocument;
			//Make sure that we have shapes to iterate over

			foreach (MSword.Range tmprng in doc.StoryRanges)
			{
				listmathML = new ArrayList();
				rng = tmprng;
				storyName = rng.StoryType.ToString();
				while (rng != null)
				{
					storyName = rng.StoryType.ToString();
					MSword.InlineShapes shapes = rng.InlineShapes;
					if (shapes != null && shapes.Count > 0)
					{
						int iCount = 1;
						int iNumShapes = 0;
						Microsoft.Office.Interop.Word.InlineShape shape;
						iNumShapes = shapes.Count;
						//iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.

						while (iCount <= iNumShapes)
						{
							if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject")
							{
								if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4")
								{
									shape = shapes[iCount];

									if (shape != null && shape.OLEFormat != null)
									{
										bool bRetVal = false;
										string strProgID;
										Guid autoConvert;
										strProgID = shape.OLEFormat.ProgID;
										bRetVal = FindAutoConvert(ref strProgID, out autoConvert);

										// if we are successful with the conversion of the CLSID we now need to query
										//  the application to see if it can actually do the work
										if (bRetVal == true)
										{
											bool bInsertable = false;
											bool bNotInsertable = false;

											bInsertable = IsCLSIDInsertable(ref autoConvert);
											bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

											//Make sure that the server of interest is insertable and not-insertable
											if (bInsertable && bNotInsertable)
											{
												bool bServerExists = false;
												string strPathToExe;
												bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

												//if the server exists then see if MathML can be retrieved for the shape
												if (bServerExists)
												{
													bool bMathML = false;
													string strVerb;
													int indexForVerb = -100;

													strVerb = "RunForConversion";

													bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
													if (bMathML)
													{
														Equation_GetMathML(ref shape, indexForVerb);
													}
												}
											}
											else
											{
												if (bInsertable != bNotInsertable)
												{
													showMsg = 1;
												}
												break;
											}
										}
									}
								}
							}
							//Increment the LCV and the number of shapes that iterated over.
							iCount++;
							iNumShapesViewed++;
						}
					}
					rng = rng.NextStoryRange;
				}
				ArrayList list = listmathML;
				multipleMathMl.Add(storyName, list);
			}
			if (showMsg == 1)
				System.Windows.Forms.MessageBox.Show("In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

			return iNumShapesViewed;
		}
		private bool FindAutoConvert(ref string ProgID, out Guid autoConvert)
		{
			bool bRetVal = false;
			Guid oGuid;
			int iCOMRetVal = 0;

			iCOMRetVal = CLSIDFromProgID(ProgID, out oGuid);

			if (iCOMRetVal == 0)
			{
				RecurseAutoConvert(ref oGuid, out autoConvert);
				bRetVal = true;
			}
			else
			{
				autoConvert = oGuid;
			}

			return bRetVal;
		}

		private void RecurseAutoConvert(ref Guid oGuid, out Guid autoConvert)
		{
			int iCOMRetVal = 0;

			iCOMRetVal = OleGetAutoConvert(ref oGuid, out autoConvert);
			if (iCOMRetVal == 0)
			{
				//If we have no error and the the CLSIDs are the same, then make sure that
				// 
				bool bGuidTheSame = false;
				try
				{
					bGuidTheSame = IsEqualGUID(ref oGuid, ref autoConvert);
				}
				catch (COMException eCOM)
				{
					MessageBox.Show(eCOM.Message);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.Message);
				}

				if (bGuidTheSame == false)
				{
					oGuid = autoConvert;
					RecurseAutoConvert(ref oGuid, out autoConvert);
				}
			}
			else
			{
				//There was some error in the auto conversion.
				// See if this guid will do the conversion.
				autoConvert = oGuid;
			}
		}

		private bool IsCLSIDInsertable(ref Guid oGuid)
		{
			bool bInsertable = false;
			//Check for the existance of the insertable key
			RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
			string strRegLocation;

			strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"Insertable";
			regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

			if (regkey != null)
				bInsertable = true;

			return bInsertable;
		}

		private bool IsCLSIDNotInsertable(ref Guid oGuid)
		{
			bool bNotInsertable = false;
			//Check for the existance of the insertable key
			RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
			string strRegLocation;
			strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"NotInsertable";

			regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

			//The not-insertable key is not present.
			if (regkey == null)
				bNotInsertable = true;

			return bNotInsertable;
		}

		private bool DoesServerExist(out string strPathToExe, ref Guid oGuid)
		{
			bool bServerExists = false;
			//Check for the existance of the insertable key
			RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
			string strRegLocation;
			strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"LocalServer32";
			regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

			if (regkey != null)
			{
				string[] valnames = regkey.GetValueNames();
				strPathToExe = "";
				try
				{
					strPathToExe = (string)regkey.GetValue(valnames[0]);
				}
				catch (Exception e)
				{
				}
				if (strPathToExe.Length > 0)
				{
					//Now check if this is a good path
					if (File.Exists(strPathToExe))
						bServerExists = true;
				}
			}
			else
			{
				strPathToExe = null;
			}

			return bServerExists;
		}

		private bool DoesServerSupportMathML(ref Guid oGuid, ref string strVerb, out int indexForVerb)
		{
			bool bIsMathMLSupported = false;
			//Check for the existance of the insertable key
			RegistryKey regkey;
			string strRegLocation;
			strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\DataFormats\GetSet";
			regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

			if (regkey != null)
			{
				string[] valnames = regkey.GetSubKeyNames();
				int x = 0;
				while (x < regkey.SubKeyCount)
				{
					RegistryKey subKey;
					if (regkey.SubKeyCount > 0)
					{
						subKey = regkey.OpenSubKey(valnames[x]);
						if (subKey != null)
						{
							string[] dataFormats = subKey.GetValueNames();
							int y = 0;
							while (y < subKey.ValueCount)
							{
								string strValue = (string)subKey.GetValue(dataFormats[y]);

								//This will accept both MathML and MathML Presentation.
								if (strValue.Contains("MathML"))
								{
									bIsMathMLSupported = true;
									break;
								}
								y++;
							}
						}
					}

					if (bIsMathMLSupported)
						break;
					x++;
				}
			}

			//Now lets check to see if the appropriate verb is supported
			if (bIsMathMLSupported)
			{
				//The return value for a verb not found will be 1000
				//
				indexForVerb = GetVerbIndex(strVerb, ref oGuid);

				if (indexForVerb == 1000)
				{
					bIsMathMLSupported = false;
				}
			}
			else
			{
				//We do not have an appropriate verb to start the server
				indexForVerb = -100;  //There is a predefined range for 
			}

			return bIsMathMLSupported;
		}

		private int GetVerbIndex(string strVerbToFind, ref Guid oGuid)
		{
			int indexForVerb = 1000;
			//Check for the existance of the insertable key
			RegistryKey regkey;
			string strRegLocation;
			strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\Verb";
			regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

			if (regkey != null)
			{
				//Lets make sure that we have some values before preceeding.
				if (regkey.SubKeyCount > 0)
				{
					int x = 0;
					int iCount = 0;

					string[] valnames = regkey.GetSubKeyNames();

					while (x < regkey.SubKeyCount)
					{
						RegistryKey subKey;
						if (regkey.SubKeyCount > 0)
						{
							subKey = regkey.OpenSubKey(valnames[x]);
							if (subKey != null)
							{
								int y = 0;
								string[] verbs = subKey.GetValueNames();
								iCount = subKey.ValueCount;
								string verb;

								//Search all of the verbs for requested string.
								while (y < iCount)
								{
									verb = (string)subKey.GetValue(verbs[y]);
									if (verb.Contains(strVerbToFind) == true)
									{
										string numVerb;
										numVerb = valnames[x].ToString();
										indexForVerb = int.Parse(numVerb);
										break;
									}
									y++;
								}
							}
						}

						//If the verb is not 1000 then break out of the verb
						if (indexForVerb != 1000)
							break;

						x++;
					}
				}
			}


			return indexForVerb;
		}

		private void Equation_GetMathML(ref Microsoft.Office.Interop.Word.InlineShape shape, int indexForVerb)
		{
			if (shape != null)
			{
				object dataObject = null;
				object objVerb;

				objVerb = indexForVerb;

				//Start MathType, and get the dataobject that is connected to the server.    
				shape.OLEFormat.DoVerb(ref objVerb);

				try
				{
					dataObject = shape.OLEFormat.Object;
				}
				catch (Exception e)
				{
					//we have an issue with trying to get the verb,
					//  There will be a attempt at another way to start the application.
					MessageBox.Show(e.Message);
				}

				IOleObject oleObject = null;

				//This is a C# version of a QueryInterface
				if (dataObject != null)
				{
					mDataObject = dataObject as IConnectDataObject;


					oleObject = dataObject as IOleObject;
				}
				else
				{
					//There was an issue with the addin trying to start with the verb we
					// knew.  A backup is to call the with the primary verb and start the 
					//  application normally.
					objVerb = MSword.WdOLEVerb.wdOLEVerbPrimary;
					shape.OLEFormat.DoVerb(ref objVerb);

					dataObject = shape.OLEFormat.Object;
					mDataObject = dataObject as IConnectDataObject;
					oleObject = dataObject as IOleObject;
				}
				//Create instances of FORMATETC and STGMEDIUM for use with IDataObject
				ConnectFORMATETC oFormatEtc = new ConnectFORMATETC();
				ConnectSTGMEDIUM oStgMedium = new ConnectSTGMEDIUM();
				DataFormats.Format oFormat;



				//Find within the clipboard system the registered clipboard format for MathML
				oFormat = DataFormats.GetFormat("MathML");

				if (mDataObject != null)
				{
					int iRetVal = 0;

					//Initialize a FORMATETC structure to get the requested data
					oFormatEtc.cfFormat = (Int16)oFormat.Id;
					oFormatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
					oFormatEtc.lindex = -1;
					oFormatEtc.ptd = (IntPtr)0;
					oFormatEtc.tymed = TYMED.TYMED_HGLOBAL;

					iRetVal = mDataObject.QueryGetData(ref oFormatEtc);
					//iRetVal will be zero if the MathML type is contained within the server.
					if (iRetVal == 0)
					{
						oStgMedium.tymed = TYMED.TYMED_NULL;
					}

					try
					{
						mDataObject.GetData(ref oFormatEtc, out oStgMedium);
					}
					catch (System.Runtime.InteropServices.COMException e)
					{
						System.Windows.Forms.MessageBox.Show(e.ToString());
						throw;
					}

					// Because we explicitly requested a MathML, we know that it is TYMED_HGLOBAL
					// lets deal with the memory here.
					if (oStgMedium.tymed == TYMED.TYMED_HGLOBAL &&
						oStgMedium.unionmember != null)
					{
						WriteOutMathMLFromStgMedium(ref oStgMedium);

						if (oleObject != null)
						{
							uint close = (uint)OLECLOSE.OLECLOSE_NOSAVE;
							// uint close = (uint)Microsoft.VisualStudio.OLE.Interop.OLECLOSE.OLECLOSE_NOSAVE;
							oleObject.Close(close);
						}
					}
				}
			}
		}

		private void WriteOutMathMLFromStgMedium(ref ConnectSTGMEDIUM oStgMedium)
		{
			IntPtr ptr;
			byte[] rawArray = null;


			//Verify that our data contained within the STGMEDIUM is non-null
			if (oStgMedium.unionmember != null)
			{
				//Get the pointer to the data that is contained
				//  within the STGMEDIUM
				ptr = oStgMedium.unionmember;

				//The pointer now becomes a Handle reference.
				HandleRef handleRef = new HandleRef(null, ptr);

				try
				{
					//Lock in the handle to get the pointer to the data
					IntPtr ptr1 = GlobalLock(handleRef);

					//Get the size of the memory block
					int length = GlobalSize(handleRef);

					//New an array of bytes and Marshal the data across.
					rawArray = new byte[length];
					Marshal.Copy(ptr1, rawArray, 0, length);

					// I will now display the text.  Create a string from the rawArray
					string str = Encoding.ASCII.GetString(rawArray);
					str = str.Substring(str.IndexOf("<mml:math"), str.IndexOf("</mml:math>") - str.IndexOf("<mml:math"));
					str = str + "</mml:math>";
					str = str.Replace("xmlns:mml='http://www.w3.org/1998/Math/MathML'", "");

					listmathML.Add(str);
				}
				catch (Exception exp)
				{
					System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + exp.ToString());
				}
				finally
				{
					//This gets called regardless within a try catch.
					//  It is a good place to clean up like this.
					GlobalUnlock(handleRef);
				}
			}
		}

		private void MTGetEquationAddinNew(String newName)
		{
			Object fileName = newName;
			object addToRecentFiles = false;
			object readOnly = false;
			object isVisible = false;
			object missing = Type.Missing;

			Microsoft.Office.Interop.Word.Document doc = this.applicationObject.Documents.Open(ref fileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
			if (doc != null)
			{
				int iNumShapesIterated = 0;
				iNumShapesIterated = IterateShapesNew(doc);
			}
		}

		private int IterateShapesNew(MSword.Document doc)
		{
			Int16 showMsg = 0;
			int iNumShapesViewed = 0;
			MSword.Range rng;
			String storyName = "";
			object missing = Type.Missing;
			object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
			object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;

			foreach (MSword.Range tmprng in doc.StoryRanges)
			{
				listmathML = new ArrayList();
				rng = tmprng;
				storyName = rng.StoryType.ToString();
				while (rng != null)
				{
					MSword.InlineShapes shapes = rng.InlineShapes;
					if (shapes != null && shapes.Count > 0)
					{
						int iCount = 1;
						int iNumShapes = 0;
						Microsoft.Office.Interop.Word.InlineShape shape;
						iNumShapes = shapes.Count;
						//iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.
						while (iCount <= iNumShapes)
						{
							if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject")
							{
								if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4")
								{
									shape = shapes[iCount];

									if (shape != null && shape.OLEFormat != null)
									{
										bool bRetVal = false;
										string strProgID;
										Guid autoConvert;
										strProgID = shape.OLEFormat.ProgID;

										bRetVal = FindAutoConvert(ref strProgID, out autoConvert);
										// if we are successful with the conversion of the CLSID we now need to query
										//  the application to see if it can actually do the work
										if (bRetVal == true)
										{
											bool bInsertable = false;
											bool bNotInsertable = false;

											bInsertable = IsCLSIDInsertable(ref autoConvert);
											bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

											//Make sure that the server of interest is insertable and not-insertable
											if (bInsertable && bNotInsertable)
											{
												bool bServerExists = false;
												string strPathToExe;
												bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

												//if the server exists then see if MathML can be retrieved for the shape
												if (bServerExists)
												{
													bool bMathML = false;
													string strVerb;
													int indexForVerb = -100;

													strVerb = "RunForConversion";

													bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
													if (bMathML)
													{
														Equation_GetMathML(ref shape, indexForVerb);
													}
												}
											}
											else
											{
												if (bInsertable != bNotInsertable)
												{
													showMsg = 1;
												}
												break;
											}
										}
									}
								}
							}
							//Increment the LCV and the number of shapes that iterated over.
							iCount++;
							iNumShapesViewed++;
						}
					}
					rng = rng.NextStoryRange;
				}
				multipleOwnMathMl.Add(storyName, listmathML);
			}

			doc.Close(ref saveChanges, ref originalFormat, ref missing);

			if (showMsg == 1)
				System.Windows.Forms.MessageBox.Show("In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

			return iNumShapesViewed;
		}


		#endregion

		#region Subdocument

		public String MasterSubDecision(string tempInput)
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
				DialogResult dr = MessageBox.Show("Do you want to translate the current document along with sub documents?", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

				if (dr == DialogResult.Yes)
					masterSubFlag = "Yes";
				else
					masterSubFlag = "No";
			}
			else
				masterSubFlag = "NoMasterSub";

			return masterSubFlag;
		}

		/*public void OoxToDaisyOwn(String docxFile, String tempInputFile, CommandBarButton ctrl, String masterSubFlag, Initialize inz)
		{
			//try
			//{
			ArrayList subList = new ArrayList();
			ArrayList subFile = new ArrayList();
			langMergeDoc = new ArrayList();
			notTranslatedDoc = new ArrayList();
			int subCount = 1;
			subList.Add(tempInputFile + "|Master");
			//OPening the Package of the Current Document
			Package packDoc;
			packDoc = Package.Open(tempInputFile, FileMode.Open, FileAccess.ReadWrite);

			//Searching for Document.xml
			foreach (PackageRelationship searchRelation in packDoc.GetRelationshipsByType(wordRelationshipType))
			{
				relationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			PackagePart mainPartxml = packDoc.GetPart(partUri);

			foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
			{
				relationship = searchRelation;
				if (relationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
				{
					if (relationship.TargetMode.ToString() == "External")
					{
						String fileName = relationship.TargetUri.ToString();

						//Checking whether sub document is of type Docx or Doc 
						if (Path.GetExtension(fileName) == ".docx" || Path.GetExtension(fileName) == ".doc")
						{
							subCount++;

							//Making a list of all Sub documents in the current Document
							if (fileName.Contains("file") && fileName.Contains("/Local Settings/Temp/"))
							{
								fileName = fileName.Replace("file:///", "");
								int indx = fileName.LastIndexOf("/Local Settings/Temp/");
								fileName = fileName.Substring(indx + 21);
								fileName = Path.GetDirectoryName(docxFile) + "//" + fileName;
								if (File.Exists(fileName))
								{
									subList.Add(fileName + "|" + relationship.Id.ToString());
									subFile.Add(fileName);
								}
							}
							else if (fileName.Contains("file"))
							{
								fileName = fileName.Replace("file:///", "");
								if (File.Exists(fileName))
								{
									subList.Add(fileName + "|" + relationship.Id.ToString());
									subFile.Add(fileName);
								}
							}
							else
							{
								fileName = Path.GetDirectoryName(docxFile) + "\\" + Path.GetFileName(fileName);
								if (File.Exists(fileName))
								{
									subList.Add(fileName + "|" + relationship.Id.ToString());
									subFile.Add(fileName);
								}
							}
						}
						//If sub document is of type other than Docx and Doc format
						else
						{
							if (fileName.Contains("file") && fileName.Contains("/Local Settings/Temp/"))
							{
								fileName = fileName.Replace("file:///", "");
								int indx = fileName.LastIndexOf("/Local Settings/Temp/");
								fileName = fileName.Substring(indx + 21);
								fileName = Path.GetDirectoryName(docxFile) + "//" + fileName;
								if (File.Exists(fileName))
								{
									notTranslatedDoc.Add(fileName);
								}
							}
							else if (fileName.Contains("file"))
							{
								fileName = fileName.Replace("file:///", "");
								if (File.Exists(fileName))
								{
									notTranslatedDoc.Add(fileName);
								}
							}
							else
							{
								fileName = Path.GetDirectoryName(docxFile) + "\\" + Path.GetFileName(fileName);
								if (File.Exists(fileName))
								{
									notTranslatedDoc.Add(fileName);
								}
							}

						}

					}
				}
			}
			packDoc.Close();

			String resultOpenSub = CheckFileOPen(subList);

			//Checking whether any Subdocumets is already Open or not
			if (resultOpenSub == "notopen")
			{
				String resultSub = CheckingSubDocs(subList);

				//Checking whether Sub documents are Simple documents or a Master document
				if (resultSub == "simple")
				{
					if (subCount == subList.Count)
					{
						saveasshapes(subFile, masterSubFlag);
						SaveasImages(subFile, masterSubFlag);
						MathMLMultiple(subList);
						if (masterSubFlag != "Yes")
							this.applicationObject.ActiveDocument.Save();
						inz.Close();
						OoxToDaisyUI(docxFile, docFile, tempInputFile, ctrl, masterSubFlag);
					}
					else
					{
						inz.Close();
						MessageBox.Show("Some Problem in Sub documents", "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
					}
				}
				else
				{
					inz.Close();
					System.Windows.Forms.MessageBox.Show("Some of the added documents are MasterSub documents.Please add simple documents.", "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
				}
			}
			else
			{
				inz.Close();
				MessageBox.Show("Some Sub documents are in open state. Please close all the Sub documents before Translation:", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}*/

		#endregion

		#region Language

		//public void GetlanguageSettings()
		//{
		//    int fileIndex;
		//    MSword.Document currentDoc = this.applicationObject.ActiveDocument;
		//    if (!currentDoc.Saved || currentDoc.FullName.LastIndexOf('.') < 0)
		//    {
		//        System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforelanguage"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
		//    }
		//    else if (currentDoc.Saved)
		//    {
		//        fileIndex = currentDoc.FullName.LastIndexOf('.');
		//        String substr = currentDoc.FullName.Substring(fileIndex);

		//        if (substr.ToLower() != ".docx")
		//        {
		//            System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
		//        }
		//        else
		//        {

		//            using (progressBar pb = new progressBar(this.applicationObject.ActiveDocument))
		//            {
		//                DialogResult dr = pb.ShowDialog();
		//                if (dr == System.Windows.Forms.DialogResult.OK)
		//                {
		//                    MessageBox.Show("Languages Updated Successfully", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
		//                }
		//            }
		//        }
		//    }

		//}


		public void GetlanguageSettings()
		{
			try
			{
				int fileIndex;
				MSword.Document currentDoc = this.applicationObject.ActiveDocument;
				if (!currentDoc.Saved || currentDoc.FullName.LastIndexOf('.') < 0)
				{
					System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforelanguage"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
				}
				else if (currentDoc.Saved)
				{
					fileIndex = currentDoc.FullName.LastIndexOf('.');
					String substr = currentDoc.FullName.Substring(fileIndex);

					if (substr.ToLower() != ".docx")
					{
						MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
					}
					else
					{
						if (currentDoc.Application.Selection.Start != currentDoc.Application.Selection.End)
						{
							ArrayList list = new ArrayList();
							foreach (string item in Enum.GetNames(typeof(MSword.WdLanguageID)))
							{
								MSword.WdLanguageID value = (MSword.WdLanguageID)Enum.Parse(typeof(MSword.WdLanguageID), item);
								list.Add(value);
							}

							Language lng = new Language(list, currentDoc.Application.Selection.Range, currentDoc);
							lng.ShowDialog();
						}
						else
						{
							MessageBox.Show("Select a paragraph for setting the Language", "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
						}
					}
				}
			}
			catch (Exception x)
			{
				MessageBox.Show(x.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		#endregion


	}

	[ComImport]
	[Guid("00000112-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IOleObject
	{
		void DoNotCall_1(object x);
		void DoNotCall_2(ref object x);
		void SetHostNames(object szContainerApp, object szContainerObj);
		void Close(uint dwSaveOption);
	};

	public enum OLECLOSE
	{
		OLECLOSE_SAVEIFDIRTY = 0,
		OLECLOSE_NOSAVE = 1,
		OLECLOSE_PROMPTSAVE = 2,
	}

}