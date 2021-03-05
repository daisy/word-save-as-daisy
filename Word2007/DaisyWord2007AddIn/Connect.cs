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


using Microsoft.Office.Interop.Word;
using Daisy.DaisyConverter.DaisyConverterLib.Converters;

namespace DaisyWord2007AddIn {
    using System;
    using System.IO;
    using System.Text;
    using System.Xml;
    using Extensibility;
    using System.Reflection;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO.Packaging;
    using System.Windows.Forms;
    using Microsoft.Office.Core;
    using System.Runtime.InteropServices;
    using MSword = Microsoft.Office.Interop.Word;
    using Daisy.DaisyConverter.DaisyConverterLib;
    using System.Xml.XPath;
    using System.Xml.Xsl;
    using System.Drawing;
    using System.Drawing.Imaging;

    using Microsoft.Win32;
    using Word = Microsoft.Office.Interop.Word.InlineShape;
    using IConnectDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
    using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
    using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
    using ConnectIEnumETC = System.Runtime.InteropServices.ComTypes.IEnumFORMATETC;
    using COMException = System.Runtime.InteropServices.COMException;
    using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;
    using System.Globalization;
    using System.Text.RegularExpressions;


    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the DaisyWord2007AddInSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />
    [GuidAttribute("18CBC6A8-BCB5-45ED-8FF3-655599E1F872"), ProgId("DaisyWord2007AddIn.Connect")]
    public class Connect : Object, IDTExtensibility2, IRibbonExtensibility {

        const string coreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        const string coreNamespace = "http://purl.org/dc/elements/1.1/";
        const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        const string subDocRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";
        const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        PackageRelationship packRelationship = null;
        XmlDocument docTemp, mergeXmlDoc, currentDocXml, validation_xml;
        String docFile, path_For_Xml, templateName, path_For_Pipeline;
        public ArrayList docValidation = new ArrayList();
        private IRibbonUI daisyRibbon;
        private bool showValidateTabBool = false;
        ArrayList mergeDoclanguage, openSubdocs;
        frmValidate2007 validateInput;
        ArrayList listmathML;
      
        ToolStripMenuItem PipelineMenuItem = null;
        private ArrayList footntRefrns;
        Hashtable multipleMathMl = new Hashtable();
        Hashtable multipleOwnMathMl = new Hashtable();
        ArrayList langMergeDoc, notTranslatedDoc;
        PackageRelationship relationship = null;


        Pipeline pipelineConversionScripts = null;
        public Pipeline PipelineConversionScripts {
            get {
                if (pipelineConversionScripts == null)
                    pipelineConversionScripts = new Pipeline(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms\scripts");
                return pipelineConversionScripts;
            }
            set => pipelineConversionScripts = value;
        }

        FileInfo postprocessScriptFile;

        /// <summary>
        ///	Implements the constructor for the Add-in object.
        ///	Place your initialization code within this method.
        /// </summary>
        public Connect() {
            try {
                this.addinLib = new Daisy.DaisyConverter.Word.Addin();
            } catch (Exception e) {
                MessageBox.Show(e.Message);
                //throw;
            }
            
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
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom) {
            try {
                this.applicationObject = (MSword.Application)application;
                this.applicationObject.DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(applicationObject_DocumentBeforeSave);
                this.applicationObject.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(applicationObject_DocumentOpen);
                this.applicationObject.DocumentChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(applicationObject_DocumentChange);
                this.applicationObject.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(applicationObject_DocumentBeforeClose);
                this.applicationObject.WindowDeactivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowDeactivateEventHandler(applicationObject_WindowDeactivate);
            } catch (Exception e) {
                MessageBox.Show(e.Message);
                //throw;
            }
        }
        /// <summary>
        /// Function to do changes in Look and Feel of UI
        /// </summary>
        void applicationObject_DocumentChange() {
            if(this.applicationObject.Documents.Count > 0) {
                MSword.Document currentDoc = this.applicationObject.ActiveDocument;
                templateName = (currentDoc.get_AttachedTemplate() as MSword.Template).Name;
                CheckforAttchments(currentDoc);
            }
            showValidateTabBool = false;
            if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");

        }
        //Envent handling deactivation of word document
        void applicationObject_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn) {

        }
        /// <summary>
        ///Core Function to check the Daisy Styles in the current Document
        /// </summary>
        /// <param name="Doc">Current Document</param>
        void applicationObject_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc) {
            //MSword.Document doc = this.applicationObject.ActiveDocument;
            templateName = (Doc.get_AttachedTemplate() as MSword.Template).Name;
            CheckforAttchments(Doc);
        }


        /// <summary>
        /// Function to check whether the Daisy Styles are new or not
        /// </summary>
        public void CheckforAttchments(MSword.Document doc) {
            //MSword.Document doc = this.applicationObject.ActiveDocument;
            string templateName = (doc.get_AttachedTemplate() as MSword.Template).Name;
            Dictionary<int, string> objArray = new Dictionary<int, string>();

            for (int iCountStyles = 1; iCountStyles <= doc.Styles.Count; iCountStyles++) {
                object ActualVal = iCountStyles;
                string strValue = doc.Styles.get_Item(ref ActualVal).NameLocal.ToString();
                if (strValue.EndsWith("(DAISY)")) {
                    objArray.Add(iCountStyles, strValue);
                }
            }
            if (objArray.Count == AddInHelper.DaisyStylesCount) {
                showViewTabBool = true;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("Button7");
            } else {
                showViewTabBool = false;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("Button7");
            }
        }

        /// <summary>
        /// Function to remove Unwanted bookmarks befreo saving the document
        /// </summary>
        /// <param name="Doc"></param>
        /// <param name="SaveAsUI"></param>
        /// <param name="Cancel"></param>
        void applicationObject_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel) {

            CustomXMLParts xmlparts = Doc.CustomXMLParts;
            XmlDocument doc = new XmlDocument();

            for (int p = 1; p <= xmlparts.Count; p++) {
                if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                    String partXml = xmlparts[p].XML;
                    doc.LoadXml(partXml);

                    foreach (object item in Doc.Bookmarks) {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                            String name = ((MSword.Bookmark)item).Name;
                            NameTable nt = new NameTable();
                            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                            nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                            XmlNodeList node = doc.SelectNodes("//a:Item[@AbbreviationName='" + name + "']", nsManager);

                            if (node.Count != 0) {
                                if (node.Item(0).Attributes[2].Value != ((MSword.Bookmark)item).Range.Text.Trim()) {
                                    object val = name;
                                    Doc.Bookmarks.get_Item(ref val).Delete();
                                }
                            }
                        } else if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                            String bkmrkName = ((MSword.Bookmark)item).Name;
                            NameTable nTable = new NameTable();
                            XmlNamespaceManager nsManager = new XmlNamespaceManager(nTable);
                            nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                            XmlNodeList node = doc.SelectNodes("//a:Item[@AcronymName='" + bkmrkName + "']", nsManager);

                            if (node.Count == 0) {
                                object val = bkmrkName;
                                Doc.Bookmarks.get_Item(ref val).Delete();
                            }
                            if (node.Count != 0) {
                                if (node.Item(0).Attributes[2].Value != ((MSword.Bookmark)item).Range.Text.Trim()) {
                                    object val = bkmrkName;
                                    Doc.Bookmarks.get_Item(ref val).Delete();
                                }
                            }
                        }
                    }
                }
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

        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom) {
            this.applicationObject = null;
            this.addinLib = null;

        }
        void applicationObject_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel) {
        }
        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom) {

        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref System.Array custom) {
            //MessageBox.Show("Starting addin");
        }

        /// <summary>
        ///      Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application is being unloaded.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref System.Array custom) {

        }

        /*Function which returns XML file to get Ribbon in Word2007*/
        string IRibbonExtensibility.GetCustomUI(string RibbonID) {
            path_For_Pipeline = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms";
            if (Directory.Exists(path_For_Pipeline)) {
                // Removing the validator button for office 2010 and later
                // For office 2007, the old validation step remains necessary
                if (this.applicationObject.Version == "12.0") {
                    return GetResource("customUI2007.xml");
                } else return GetResource("customUI.xml");
            } else {
                return GetResource("customUIOld.xml");
            }
        }

        /*Function to get data of a file */
        private string GetResource(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames()) {
                if (name.EndsWith(resourceName)) {
                    System.IO.TextReader tr = new System.IO.StreamReader(asm.GetManifestResourceStream(name));
                    //Debug.Assert(tr != null);
                    string resource = tr.ReadToEnd();

                    tr.Close();
                    return resource;
                }
            }
            return null;
        }


        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage(IRibbonControl control) {
            return this.addinLib.GetLogo("speaker.jpg");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImageSingleXML(IRibbonControl control) {
            return this.addinLib.GetLogo("Singlexml.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImageMultipleXML(IRibbonControl control) {
            return this.addinLib.GetLogo("Multiplexml.png");
        }


        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage1(IRibbonControl control) {
            return this.addinLib.GetLogo("ABBR.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage2(IRibbonControl control) {
            return this.addinLib.GetLogo("ABBR2.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage3(IRibbonControl control) {
            return this.addinLib.GetLogo("ACR.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage4(IRibbonControl control) {
            return this.addinLib.GetLogo("ACR2.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage5(IRibbonControl control) {
            return this.addinLib.GetLogo("help.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage6(IRibbonControl control) {
            return this.addinLib.GetLogo("DaisyLogo.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage7(IRibbonControl control) {
            return this.addinLib.GetLogo("subfolder.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage8(IRibbonControl control) {
            return this.addinLib.GetLogo("version.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage9(IRibbonControl control) {
            return this.addinLib.GetLogo("validate.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage10(IRibbonControl control) {
            return this.addinLib.GetLogo("import.png");
        }

        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage11(IRibbonControl control) {
            return this.addinLib.GetLogo("footnotes.png");
        }


        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage12(IRibbonControl control) {
            return this.addinLib.GetLogo("gear.png");
        }


        /*Function to get image before the label*/
        public stdole.IPictureDisp GetImage13(IRibbonControl control) {
            return this.addinLib.GetLogo("Language.png");
        }


        /*Function to get label in the Word2007 Ribbon*/
        public string GetLabel(IRibbonControl control) {
            return this.addinLib.GetString(control.Id + "Label");
        }
        /*Function to get the description for a label*/
        public string GetDescription(IRibbonControl control) {
            return this.addinLib.GetString(control.Id + "Description");
        }

        public void OnLoad(IRibbonUI ribbon) {
            daisyRibbon = ribbon;
        }
        private bool showViewTabBool = false;


        public void GetDaisySettings(IRibbonControl control) {
            DAISY_Settings daisyfrm = new DAISY_Settings();
            daisyfrm.ShowDialog();
        }


        /// <summary>
        /// Function to update the Daisy Styles
        /// </summary>
        /// <param name="control"></param>
        public void Update_Styles(IRibbonControl control) {
            object copyTempName = null;
            MSword.Document docTemplate = new Microsoft.Office.Interop.Word.Document();
            MSword.Document docActive = this.applicationObject.ActiveDocument;
            object missing = Type.Missing;
            try {
                string templateName = (docActive.get_AttachedTemplate() as MSword.Template).Name;
                //TODO: rename template
                string templatePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "template2007.dotx");
                object newTemp = false;
                object documentType = Microsoft.Office.Interop.Word.WdNewDocumentType.wdNewBlankDocument;
                object visble = false;

                copyTempName = Path.GetTempFileName() + ".dotx";

                if (!File.Exists(copyTempName.ToString())) {
                    File.Copy(templatePath, copyTempName.ToString());
                }

                object tempPath = (object)templatePath;
                docTemplate = applicationObject.Documents.Add(ref tempPath, ref newTemp, ref documentType, ref visble);


                foreach (MSword.Style styleObj in docTemplate.Styles) {
                    if (styleObj.NameLocal.EndsWith("(DAISY)")) {
                        this.applicationObject.OrganizerCopy(copyTempName.ToString(), docActive.FullName, styleObj.NameLocal, MSword.WdOrganizerObject.wdOrganizerObjectStyles);
                    }
                }
                object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                docTemplate.Close(ref saveChanges, ref missing, ref missing);

                this.applicationObject.NormalTemplate.Save();

                if (File.Exists(copyTempName.ToString())) {
                    File.Delete(copyTempName.ToString());
                }

                showViewTabBool = true;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("Button7");

            } catch (Exception ex) {
                string stre = ex.Message;
                MessageBox.Show(ex.Message.ToString(), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                docTemplate.Close(ref saveChanges, ref missing, ref missing);
            }

        }

        public bool getEnabled(IRibbonControl control) {
            return !showViewTabBool;
        }
        //Method for Disabling validate tab
        public bool buttonValidatePressed(IRibbonControl control) {
            return showValidateTabBool;
        }
        //Method for Enabling validate tab
        public bool getValidateEnabled(IRibbonControl control) {
            return !showValidateTabBool;
        }

        /*Function validate input docx file against DAISY rule*/
        public void Validate(IRibbonControl control, bool flip) {

            float officeVersion = float.Parse(this.applicationObject.Version, CultureInfo.InvariantCulture.NumberFormat);
            if (officeVersion >= 14.0) {
                //MessageBox.Show("For office 2010 and superior, the word validation is being replaced for the Microsoft Accessibility Checker.\r\n" +
                //"You may find it in the 'File' > 'Info' > 'Check for issues' menu, or under the Review tab of the ribbon for Office 2016 and newer.\r\n" +
                //"Do you want to open it now ?", "Deprecated functionnality", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if(this.applicationObject.Documents.Count > 0) {
                    // using sendkeys to access the pane, as it is not exposed through any API i tested
                    // i tried launching it via command bars and the taskpanes interfaces but it was not exposed through those
                    if (officeVersion >= 16.0) SendKeys.SendWait("%ra1"); // open the checker using the ribbon button added in 2016
                    else if (officeVersion == 15.0) SendKeys.SendWait("%fxxa"); // open the checker through the File > info > Check for issues > check accessibility button
                    else SendKeys.SendWait("%fxta"); // open the checker through the File > info > Check for issues > check accessibility button
                }
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
            } else try {
                docValidation = new ArrayList();
                int ind;
                MSword.Document doc = this.applicationObject.ActiveDocument;
                //Path of the active document
                string path = doc.FullName;
                //Checking whether document is saved
                if (!doc.Saved || doc.FullName.LastIndexOf('.') < 0) {
                    System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforeValidate"), "SaveAsDAISY",
                                                         System.Windows.Forms.MessageBoxButtons.OK,
                                                         System.Windows.Forms.MessageBoxIcon.Stop);
                    showValidateTabBool = false;
                    if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
                }
                //if saved
                else if (doc.Saved) {
                    ind = doc.FullName.LastIndexOf('.');
                    String substr = doc.FullName.Substring(ind);
                    //Checking the saved format is docx
                    if (substr.ToLower() != ".docx") {
                        System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY",
                                                             System.Windows.Forms.MessageBoxButtons.OK,
                                                             System.Windows.Forms.MessageBoxIcon.Stop);
                        showValidateTabBool = false;
                        if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
                        return;
                    }
                    //if docx format
                    else {
                        docFile = this.addinLib.GetTempPath((string)doc.FullName, ".docx");
                        File.Copy((string)doc.FullName, (string)docFile);
                    }
                    DeleteBookMark();

                    object saveChangesWord = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                    object originalFormatWord = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
                    object FileName = (object)path;
                    object AddToRecentFiles = false;
                    object routeDocument = Type.Missing;
                    //object Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
                    //Getting validation xml from executing assembly
                    path_For_Xml = AddInHelper.AppDataSaveAsDAISYDirectory;

                    validation_xml = new XmlDocument();
                    //chcking for validation xml
                    if (File.Exists(path_For_Xml + "\\prepopulated_validationVer21.xml")) {
                        validation_xml.Load(path_For_Xml + "\\prepopulated_validationVer21.xml");

                        //Opening package
                        Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);

                        foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                            packRelationship = searchRelation;
                            break;
                        }
                        Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                        PackagePart mainPartxml = pack.GetPart(partUri);
                        currentDocXml = new XmlDocument();
                        currentDocXml.Load(mainPartxml.GetStream());
                        pack.Close();
                        //Progress bar
                        using (ProgressValidation prog = new ProgressValidation(currentDocXml, validation_xml)) {
                            //Progress bar dialog
                            DialogResult dr = prog.ShowDialog();
                            //If work complete
                            if (dr == System.Windows.Forms.DialogResult.OK) {
                                docValidation = prog.GetFileNames;
                                //If docValidation array list contains validation error
                                if (docValidation.Count != 0) {
                                    MessageBox.Show("This document contains validation errors", "Validation completed",
                                                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    //Closing the word document
                                    this.applicationObject.ActiveDocument.Close(ref saveChangesWord, ref originalFormatWord, ref routeDocument);
                                    //Opening the package of the closed document
                                    Package packClose = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
                                    foreach (PackageRelationship searchRelation in packClose.GetRelationshipsByType(wordRelationshipType)) {
                                        packRelationship = searchRelation;
                                        break;
                                    }
                                    Uri packClosePartUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                                    PackagePart packCloseMainPartxml = packClose.GetPart(packClosePartUri);
                                    currentDocXml = prog.XmlFileNames;
                                    using (Stream str = packCloseMainPartxml.GetStream(FileMode.Create)) {
                                        using (StreamWriter ts = new StreamWriter(str)) {
                                            ts.Write(currentDocXml.OuterXml);
                                            ts.Flush();
                                            ts.Close();
                                        }
                                    }
                                    packClose.Flush();
                                    packClose.Close();
                                    object formatWord = MSword.WdSaveFormat.wdFormatXMLDocument;
                                    object visible = true;
                                    object repair = false;
                                    object readOnly = false;
                                    //Opening the word document
                                    this.applicationObject.Documents.Open(ref FileName, ref routeDocument, ref readOnly, ref AddToRecentFiles,
                                                                          ref routeDocument, ref routeDocument, ref routeDocument,
                                                                          ref routeDocument, ref routeDocument, ref routeDocument,
                                                                          ref routeDocument, ref visible, ref routeDocument, ref routeDocument,
                                                                          ref routeDocument, ref routeDocument);
                                    //Initializing frmValidate2007 form
                                    validateInput = new frmValidate2007();
                                    //Calling method and passing validation xml and active document
                                    validateInput.SetApp(docValidation, this.applicationObject.ActiveDocument);
                                    //Showing frmValidate2007 form
                                    validateInput.Show(new WordWin32Window(this.applicationObject));
                                    //Event on close of frmValidate2007 form
                                    validateInput.FormClosed += new FormClosedEventHandler(validateInput_FormClosed);
                                    //Disabling validation tab
                                    showValidateTabBool = flip;
                                    daisyRibbon.Invalidate();
                                } else {
                                    //Message if document is valid
                                    MessageBox.Show("This document is valid", "Validation completed", System.Windows.Forms.MessageBoxButtons.OK,
                                                    System.Windows.Forms.MessageBoxIcon.Information);
                                    //Enabling Validate tab
                                    showValidateTabBool = false;
                                    if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
                                }
                            }
                            //Progress bar canceled
                            if (dr == System.Windows.Forms.DialogResult.Cancel) {
                                MessageBox.Show("Validation stopped", "Quit validation", System.Windows.Forms.MessageBoxButtons.OK,
                                                System.Windows.Forms.MessageBoxIcon.Stop);
                                showValidateTabBool = false;
                                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
                            }

                        }

                    } else {
                        MessageBox.Show("Validation XML does not exits");
                        AddinLogger.Error(string.Format("Validation XML does not exits : {0}", path_For_Xml + "\\prepopulated_validationVer21.xml"));
                    }
                }
            } catch (Exception ex) {
                AddinLogger.Error(ex);
                MessageBox.Show(ex.Message, "Validation error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        /*Function which deletes bookmarks on close of validateform */
        void validateInput_FormClosed(object sender, FormClosedEventArgs e) {
            try {
                DeleteBookMark();
                showValidateTabBool = false;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidate");
                this.applicationObject.ActiveDocument.Save();
            } catch {

            }
        }

        /// <summary>
        /// Function to generate Random ID
        /// </summary>
        /// <returns></returns>
        public long GenerateId() {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }


        /*Function which deletes bookmark */
        public void DeleteBookMark() {
            try {
                MSword.Document doc = this.applicationObject.ActiveDocument;
                //Looping through the document for each Bookmark
                foreach (object item in doc.Bookmarks) {
                    //Checking for bookmark Rule
                    if (((MSword.Bookmark)item).Name.StartsWith("Rule", StringComparison.CurrentCulture)) {
                        object val = ((MSword.Bookmark)item).Name;
                        doc.Bookmarks.get_Item(ref val).Delete();
                    }
                }
                //Saving the active word document
                doc.Save();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Function to show Help File
        /// </summary>
        /// <param name="control"></param>
        public void HelpUI(IRibbonControl control) {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm");
        }

        public void ShowWordManual(IRibbonControl control) {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "SaveAsDAISY_Instruction_Manual_Jan_2021.docx");
        }

        public void ShowAuthoringGuidelines(IRibbonControl control) {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Authoring_Guidelines_For_using_SaveAsDAISY.docx");
        }

        public void ShowDtBookManual(IRibbonControl control) {
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Save-as-instruction-manual");
        }



        /// <summary>
        /// Function to show the current version of the SaveAsDAISY Addin
        /// </summary>
        /// <param name="control"></param>
        public void AboutUI(IRibbonControl control) {
            About abtForm = new About();
            abtForm.ShowDialog();
        }

        #region save as daisy from single docx

        public bool SaveAsSingleDaisyInQuiteMode(Document activeDocument, TranslationParametersBuilder preporator, string output) {
            string ctrlId = "DaisySingle";
            IPluginEventsHandler eventsHandler = new PluginEventsQuiteHandler();
            PreparetionResult preparetionResult = PrepareForSaveAsDaisy(eventsHandler, activeDocument, ctrlId);
            if (preparetionResult.IsSuccess) {
                OoxToDaisyParameters parameters = BuildParameters(preparetionResult, ctrlId);
                addinLib.OoxToDaisyWithoutUI(parameters, preporator, output, string.Empty);
                return true;
            }
            return false;
        }

        /**
         * UI Action called by single word buttons (dtbook XML or DAISY book)
         * Function which Translates docx file to Daisy format 
         */
        public void SaveAsDaisy(IRibbonControl control) {
            try {
                IPluginEventsHandler eventsHandler = new PluginEventsUIHandler();
                PreparetionResult preparetionResult = PrepareForSaveAsDaisy(eventsHandler, applicationObject.ActiveDocument, control.Tag);
                if (preparetionResult.IsSuccess) {
                    OoxToDaisyParameters parameters = BuildParameters(preparetionResult, control.Tag);
                    addinLib.StartSingleWordConversion(parameters);
                    // Removing call due to single reference
                    // StartSingleWordConversion(preparetionResult, control.Tag);
                } else if (!preparetionResult.IsCanceled) {
                    MessageBox.Show(preparetionResult.LastMessage, "Conversion stopped");
                } 
                applicationObject.ActiveDocument.Save();
            } catch (Exception e) {
                ExceptionReport report = new ExceptionReport(e);
                report.Show();
            }
            
        }

        public OoxToDaisyParameters BuildParameters(PreparetionResult preparetionResult, string ctrlId) {
            

            OoxToDaisyParameters parameters = new OoxToDaisyParameters();
            parameters.InputFile = preparetionResult.OriginalFilePath;
            parameters.TempInputFile = preparetionResult.TempFilePath;
            parameters.TempInputA = preparetionResult.TempFilePath;
            parameters.Version = this.applicationObject.Version;
            parameters.ControlName = ctrlId;
            parameters.ObjectShapes = preparetionResult.ObjectShapes;
            parameters.ListMathMl = multipleMathMl;
            parameters.ImageIds = preparetionResult.ImageId;
            parameters.InlineShapes = preparetionResult.InlineShapes;
            parameters.InlineIds = preparetionResult.InlineId;
            parameters.MasterSubFlag = preparetionResult.MasterSubFlag;
            if (!AddInHelper.buttonIsSingleWordToXMLConversion(ctrlId)) {
                parameters.ScriptPath = PipelineConversionScripts.ScriptsInfo[ctrlId].FullName;
                parameters.Directory = string.Empty;
            } else if (PipelineConversionScripts.ScriptsInfo.TryGetValue("_postprocess", out postprocessScriptFile)) {
                // Note : adding a postprocess script for dtbook pipeline special treatment
                parameters.ScriptPath = postprocessScriptFile.FullName;
                parameters.Directory = string.Empty;
            } else parameters.ScriptPath = null;
            return parameters;
        }

        /**
         * Function to retrieve conversion parameters and start conversion process from the library
         *
        public void StartSingleWordConversion(PreparetionResult preparetionResult, string ctrlId) {
            OoxToDaisyParameters parameters = BuildParameters(preparetionResult, ctrlId);
            addinLib.StartSingleWordConversion(parameters);
        }/* */


        public PreparetionResult PrepareForSaveAsDaisy(IPluginEventsHandler eventsHandler, Document activeDocument, string controlId) {
            PreparetionResult result = new PreparetionResult();
            listmathML = new ArrayList();
            multipleMathMl = new Hashtable();
            result.ObjectShapes = new ArrayList();
            result.ImageId = new ArrayList();
            result.InlineShapes = new ArrayList();
            result.InlineId = new ArrayList();
            int fileIndex;
            Document currentDoc = activeDocument;

            if (!currentDoc.Saved || currentDoc.FullName.LastIndexOf('.') < 0) {
                eventsHandler.OnStop(addinLib.GetString("DaisySaveDocumentBeforeExport"));
                return PreparetionResult.Failed("Please save your document before going further.");
            }
            
            fileIndex = currentDoc.FullName.LastIndexOf('.');
            String substr = currentDoc.FullName.Substring(fileIndex);

            if (substr.ToLower() != ".docx") {
                eventsHandler.OnStop(addinLib.GetString("DaisySaveDocumentin2007"));
                return PreparetionResult.Failed("The document is not a docx file saved on your system.");
            }
            object missing = Type.Missing;

            // Adding a filename check on the current docx to prevent problematic characters in the filename
            StringBuilder errorFileNameMessage = new StringBuilder("Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n");
            // For dtbook conversion, any correct system file name will work, except for the ones with commas in it in my tests
            // Possibly an error in the pipeline commande line parsing of arguments in the pipeline side
            string authorizedNamePattern = @"^[^,]+$"; 
            if (AddInHelper.buttonIsSingleWordToXMLConversion(controlId)) {
                errorFileNameMessage.Append("Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically.");
            } else {
                // TODO : specific name pattern following daisy book naming convention to find
                authorizedNamePattern = @"^[a-zA-Z0-9_\-\.]+\.docx$";
                errorFileNameMessage.Append(
                    "Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
                        "and underscores (_) are allowed in DAISY file names." +
                    "\r\nAny other characters (including spaces) will be replaced automaticaly by underscores.");
            }
            errorFileNameMessage.Append(
                "\r\n" +
                "\r\nDo you want to save this document under a new name ?" +
                "\r\nThe document with the original name will not be deleted." +
                "\r\n" +
                "\r\n(Click Yes to save the document under a new name and use the new one, " +
                    "No to continue with the current document, " +
                    "or Cancel to abort the conversion)");


            Regex validator = new Regex(authorizedNamePattern);
            bool nameIsValid;
            do {
                bool docIsRenamed = false;
                if (!validator.IsMatch(currentDoc.Name)) { // check only name (i assume it may still lead to problem if path has commas)
                    DialogResult userAnswer = MessageBox.Show(errorFileNameMessage.ToString(), "Unauthorized characters in the document filename", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Warning);
                    if (userAnswer == DialogResult.Yes) {
                        Dialog dlg = this.applicationObject.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFileSaveAs];
                        int saveResult = dlg.Show(ref missing);
                        if (saveResult == -1) { // ok pressed, see https://docs.microsoft.com/fr-fr/dotnet/api/microsoft.office.interop.word.dialog.show?view=word-pia#Microsoft_Office_Interop_Word_Dialog_Show_System_Object__
                            docIsRenamed = true;
                        } else return PreparetionResult.Canceled("User canceled a renaming request for an invalid docx filename");
                    } else if (userAnswer == DialogResult.Cancel) {
                        return PreparetionResult.Canceled("User canceled a renaming request for an invalid docx filename");
                    }
                    // else a sanitize path in the DaisyAddinLib will replace commas by underscore.
                    // Other illegal characters regarding the conversion to DAISY book are replaced by underscore by the pipeline itself
                    // While image names seems to be sanitized in other process
                }
                nameIsValid = !docIsRenamed;
            } while (!nameIsValid);

            result.InitializeWindow = new Initialize();
            object originalPath = currentDoc.FullName;
            object tmpFileName = this.addinLib.GetTempPath((string)originalPath, ".docx");
            object newName = Path.GetTempFileName() + Path.GetExtension((string)originalPath);
            
            // Duplicate the current doc and use the copy
            object addToRecentFiles = false;
            object readOnly = false;

            // visibility
            object visible = true;
            object invisible = false;

            object format = WdSaveFormat.wdFormatXMLDocument;

            // FIX 05/03/2021 : Error is raised here for onedrive files that are using "http(s)" urls
            // For now we replace the copy by a standard office save and reopen the original file
            //File.Copy((string)currentDoc.FullName, (string)newName);
            //newDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //try {
            //    File.Delete((string)newName);
            //} catch (IOException) {}

            // Save a copy and reopen the the original document
            currentDoc.SaveAs(ref tmpFileName, ref format, ref missing, ref missing, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            currentDoc = this.applicationObject.Documents.Open(ref originalPath);

            // Open, or retrieve the temp file if opened in word
            Document newDoc = this.applicationObject.Documents.Open(ref tmpFileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref invisible, ref missing, ref missing, ref missing, ref missing);
            // close the temp file
            object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            newDoc.Close(ref saveChanges, ref originalFormat, ref missing); 
            
            docFile = (string)tmpFileName;

            PrepopulateDaisyXml prepopulateDaisyXml = new PrepopulateDaisyXml();
            prepopulateDaisyXml.Creator = DocPropCreator();
            prepopulateDaisyXml.Title = DocPropTitle();
            prepopulateDaisyXml.Publisher = DocPropPublish();
            prepopulateDaisyXml.Save();

            // FIX 05/03/2021 : OriginalFilePath is used as InputFile in OoxToDaisyParameters. 
            // For onedrive url (starting with https) the temp copy is used as input path instead of the real one
            result.OriginalFilePath = currentDoc.FullName.StartsWith("http") ? docFile : currentDoc.FullName;
            result.TempFilePath = docFile;

            result.MasterSubFlag = MasterSubDecision(docFile, eventsHandler);
            result.InitializeWindow.Show();
            Application.DoEvents();
            try {
                saveasshapes();
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:" +
                    "\r\n- " + e.Message + 
                    "\r\n" + e.StackTrace);
            }
            try {
                SaveasImages();
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing images and may prevent the rest of the conversion to success:" +
                    "\r\n- " + e.Message +
                    "\r\n" + e.StackTrace);
            }

            
            this.applicationObject.ActiveDocument.Save();
            if (result.MasterSubFlag == "Yes") {
                result.IsSuccess = OoxToDaisyOwn(result, controlId, eventsHandler);
            } else {
                MTGetEquationAddin(eventsHandler);
                result.IsSuccess = true;
                
            }
            result.InitializeWindow.Close();

            return result;
        }

        #endregion

        private MSword.Application applicationObject;
        private DaisyAddinLib addinLib;

        #region Abbreviation and Acronyms

        /// <summary>
        /// Core Function for Managing Abbreviation
        /// </summary>
        /// <param name="control"></param>
        public void ManageAbbreviation(IRibbonControl control) {
            try {

                MSword.Document doc = this.applicationObject.ActiveDocument;
                if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection) {
                    Abbreviation form = new Abbreviation(doc, true);
                    form.ShowDialog();
                } else {
                    MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Core Function for Managing Abbreviations
        /// </summary>
        /// <param name="control"></param>
        public void ManageAcronym(IRibbonControl control) {
            try {
                MSword.Document doc = this.applicationObject.ActiveDocument;
                if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection) {
                    Abbreviation form = new Abbreviation(doc, false);
                    form.ShowDialog();
                } else {
                    MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Core function to Mark Abbreviations
        /// </summary>
        /// <param name="control"></param>
        public void MarkAsAbbreviationUI(IRibbonControl control) {
            MSword.Document doc = this.applicationObject.ActiveDocument;

            if ((this.applicationObject.Selection.Start == this.applicationObject.Selection.End) || this.applicationObject.Selection.Text.Equals("\r")) {
                MessageBox.Show(addinLib.GetString("AbbreviationText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (validateSpace(this.applicationObject.Selection.Text.Trim())) {
                MessageBox.Show(addinLib.GetString("AbbreviationproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateWord(this.applicationObject.Selection)) {
                MessageBox.Show(addinLib.GetString("AbbreviationproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateBookMark(this.applicationObject.Selection) == "Abbrtrue") {
                MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateBookMark(this.applicationObject.Selection) == "Acrtrue") {
                MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else {
                Mark mrkForm = new Mark(doc, true);
                mrkForm.ShowDialog();
            }
        }

        /// <summary>
        /// Function to check whether the selected Bookmark is Abbreviation\Acronym
        /// </summary>
        /// <param name="select">Range of text</param>
        /// <returns></returns>
        public String ValidateBookMark(MSword.Selection select) {
            MSword.Document currentDoc = this.applicationObject.ActiveDocument;
            String bkmrkValue = "Nobookmarks";

            if (this.applicationObject.Selection.Bookmarks.Count > 0) {

                foreach (object item in currentDoc.Bookmarks) {
                    if (((MSword.Bookmark)item).Range.Text.Trim() == this.applicationObject.Selection.Text.Trim()) {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
                            bkmrkValue = "Abbrtrue";
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                            bkmrkValue = "Acrtrue";
                    }
                }
            }

            return bkmrkValue;
        }

        /// <summary>
        /// Function to check the selected Text is valid for Abbreviation\Acronym or not
        /// </summary>
        /// <param name="select">Range of text</param>
        /// <returns></returns>
        public bool ValidateWord(MSword.Selection select) {
            bool flag = false;
            MSword.Document doc = this.applicationObject.ActiveDocument;

            if (select.Words.Count == 1) {
                if (this.applicationObject.Selection.Words.Last.Text.Trim() != this.applicationObject.Selection.Text.Trim())
                    flag = true;
            }
            return flag;
        }

        /// <summary>
        /// Core Function for Managing Acronyms
        /// </summary>
        /// <param name="control"></param>
        public void MarkAsAcronymUI(IRibbonControl control) {
            MSword.Document doc = this.applicationObject.ActiveDocument;

            if (this.applicationObject.Selection.Start == this.applicationObject.Selection.End || this.applicationObject.Selection.Text.Equals("\r")) {
                MessageBox.Show(addinLib.GetString("AcronymText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (validateSpace(this.applicationObject.Selection.Text.Trim())) {
                MessageBox.Show(addinLib.GetString("AcronymproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateWord(this.applicationObject.Selection)) {
                MessageBox.Show(addinLib.GetString("AcronymproperText"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateBookMark(this.applicationObject.Selection) == "Acrtrue") {
                MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if (ValidateBookMark(this.applicationObject.Selection) == "Abbrtrue") {
                MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else {
                Mark mrkForm = new Mark(doc, false);
                mrkForm.ShowDialog();
            }
        }

        /// <summary>
        /// Functio to check the selected Text is valid for Abbreviation\Acronym or not
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public bool validateSpace(String text) {
            if (text.Contains(" "))
                return true;
            else if (text == "")
                return true;
            else
                return false;
        }

        #endregion

        #region Document Properties

        /// <summary>
        /// Function to get the Title of the Current Document
        /// </summary>
        /// <returns>Title</returns>
        public String DocPropTitle() {
            int titleFlag = 0;
            String styleVal = "";
            string msgConcat = "";
            Package pack;
            pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
            String strTitle = pack.PackageProperties.Title;
            pack.Close();
            if (strTitle != "")
                return strTitle;
            else {
                pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
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
                for (int j = 0; j < getParagraph.Count; j++) {
                    XmlAttributeCollection paraGraphAttribute = getParagraph[j].Attributes;
                    for (int i = 0; i < paraGraphAttribute.Count; i++) {
                        if (paraGraphAttribute[i].Name == "w:val") {
                            styleVal = paraGraphAttribute[i].Value;
                        }
                        if (styleVal != "" && styleVal == "Title") {
                            XmlNodeList getStyle = getParagraph[j].ParentNode.ParentNode.SelectNodes("w:r", nsManager);
                            if (getStyle != null) {
                                for (int k = 0; k < getStyle.Count; k++) {
                                    XmlNode getText = getStyle[k].SelectSingleNode("w:t", nsManager);
                                    msgConcat = msgConcat + " " + getText.InnerText;
                                }
                            }
                            titleFlag = 1;
                            break;
                        }
                        if (titleFlag == 1) {
                            break;
                        }
                    }
                    if (titleFlag == 1) {
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
        public String DocPropCreator() {
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
        public String DocPropPublish() {
            Package pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(appRelationshipType)) {
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

        #endregion

        #region Multiple OOXML functions

        /// <summary>
        /// Function to Check selected Documents are Master/Sub dcouemnts or simple Documents 
        /// </summary>
        /// <param name="listSubDocs">Seleted Documents</param>
        /// <param name="value">bool value</param>
        /// <returns>document type</returns>
        //public String CheckingSubDocs(ArrayList listSubDocs)
        //{
        //    String resultSubDoc = "simple";
        //    for (int i = 0; i < listSubDocs.Count; i++)
        //    {
        //        string[] splitName = listSubDocs[i].ToString().Split('|');
        //        Package pack;
        //        pack = Package.Open(splitName[0].ToString(), FileMode.Open, FileAccess.ReadWrite);

        //        foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
        //        {
        //            packRelationship = searchRelation;
        //            break;
        //        }

        //        Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
        //        PackagePart mainPartxml = pack.GetPart(partUri);

        //        foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
        //        {
        //            packRelationship = searchRelation;
        //            if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
        //            {
        //                if (packRelationship.TargetMode.ToString() == "External")
        //                {
        //                    resultSubDoc = "complex";
        //                }
        //            }
        //        }
        //        pack.Close();
        //    }
        //    return resultSubDoc;
        //}

        /// <summary>
        /// Function to Check selected Documents are Master/Sub dcouemnts or simple Documents 
        /// </summary>
        /// <param name="listSubDocs">Seleted Documents</param>
        /// <param name="value">bool value</param>
        /// <returns>document type</returns>
        //public String CheckingSubDocs(ArrayList listSubDocs, bool value)
        //{
        //    String resultSubDoc = "simple";
        //    for (int i = 1; i < listSubDocs.Count; i++)
        //    {
        //        string[] splitName = listSubDocs[i].ToString().Split('|');
        //        Package pack;
        //        pack = Package.Open(splitName[0].ToString(), FileMode.Open, FileAccess.ReadWrite);

        //        foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType))
        //        {
        //            packRelationship = searchRelation;
        //            break;
        //        }

        //        Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
        //        PackagePart mainPartxml = pack.GetPart(partUri);

        //        foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
        //        {
        //            packRelationship = searchRelation;
        //            if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument")
        //            {
        //                if (packRelationship.TargetMode.ToString() == "External")
        //                {
        //                    resultSubDoc = "complex";
        //                }
        //            }
        //        }
        //        pack.Close();
        //    }
        //    return resultSubDoc;
        //}

        public void SignleDaisyFromMultipleInQuiteMode(ArrayList documents, string outputFilePath, TranslationParametersBuilder preporator) {
            Application.DoEvents();
            multipleMathMl = new Hashtable();

            Initialize inz = new Initialize();
            inz.Show();
            Application.DoEvents();
            ArrayList subList = documents;
            string output_Pipeline = string.Empty;
            mergeXmlDoc = new XmlDocument();
            mergeDoclanguage = new ArrayList();
            String individual_docs = "individual";
            String resultOpenSub = CheckFileOPen(subList);
            try {
                saveasshapes(subList, "No");
            } catch (Exception e) {
                //MessageBox.Show("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }
            try {
                SaveasImages(subList, "No");
            } catch (Exception e) {
                // MessageBox.Show("An error occured while preprocessing images and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }

            if (resultOpenSub == "notopen") {
                String resultSub = SubdocumentsManager.CheckingSubDocs(subList);
                if (resultSub == "simple") {
                    MathMLMultiple(subList);
                    inz.Close();
                    bool result = this.addinLib.OoxToDaisySubWithoutUI(outputFilePath, subList, individual_docs, preporator.BuildTranslationParameters(), "DaisyMultiple", multipleMathMl, output_Pipeline);
                } else {
                    inz.Close();
                    //MessageBox.Show(addinLib.GetString("AddSimpleMasterSub"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                }
            } else {
                inz.Close();
                String tempArray = "";
                for (int i = 0; i < openSubdocs.Count; i++)
                    tempArray += (i + 1) + ". " + openSubdocs[i].ToString();
                //MessageBox.Show(this.addinLib.GetString("OPenState") + "\n\n" + tempArray, "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            }
        }

        /// <summary>
        /// Core function to Translate the bunch of Documents
        /// </summary>
        /// <param name="control"></param>
        public void Mutiple(IRibbonControl control) {
            IPluginEventsHandler eventsHandler = new PluginEventsUIHandler();
            Application.DoEvents();

            MultipleSub mulsubDoc;
            if (AddInHelper.IsSingleDaisyFromMultipleButton(control.Tag)) {
                if (PipelineConversionScripts.ScriptsInfo.TryGetValue("_postprocess", out postprocessScriptFile))
                    mulsubDoc = new MultipleSub(this.addinLib.ResManager, this.applicationObject.Version, control.Tag, postprocessScriptFile.FullName, "");
                else mulsubDoc = new MultipleSub(this.addinLib.ResManager, this.applicationObject.Version, control.Tag);
            } else {
                mulsubDoc = new MultipleSub(this.addinLib.ResManager, this.applicationObject.Version, control.Tag, PipelineConversionScripts.ScriptsInfo[control.Tag].FullName, "");
            }
            int subDocFlag = mulsubDoc.DoTranslate();

            if (subDocFlag != 1) return;

            multipleMathMl = new Hashtable();
            Initialize inz = new Initialize();
            inz.Show();
            Application.DoEvents();
            ArrayList subList = mulsubDoc.GetFileNames;
            string outputFilePath = mulsubDoc.GetOutputFilePath;
            string output_Pipeline = mulsubDoc.pipeOutput;
            mergeXmlDoc = new XmlDocument();
            mergeDoclanguage = new ArrayList();
            String individual_docs = "individual";
            String resultOpenSub = CheckFileOPen(subList);
            try {
                saveasshapes(subList, "No");
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }
            try {
                SaveasImages(subList, "No");
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing images and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }



            if (resultOpenSub != "notopen") {
                inz.Close();
                String tempArray = "";
                for (int i = 0; i < openSubdocs.Count; i++)
                    tempArray += (i + 1) + ". " + openSubdocs[i].ToString();
                eventsHandler.OnStop(this.addinLib.GetString("OPenState") + "\n\n" + tempArray);
                return;
            }

            String resultSub = SubdocumentsManager.CheckingSubDocs(subList);
            if (resultSub != "simple") {
                inz.Close();
                eventsHandler.OnStop(addinLib.GetString("AddSimpleMasterSub"));
                return;
            }

            MathMLMultiple(subList);
            inz.Close();

            bool result = this.addinLib.OoxToDaisySub(outputFilePath, subList, individual_docs, mulsubDoc.HTable, control.Tag,
                                                      multipleMathMl, output_Pipeline);
            if (!(AddInHelper.IsSingleDaisyFromMultipleButton(control.Tag) && mulsubDoc.getParser == null) && result) {
                try {
                    mulsubDoc.getParser.ExecuteScript(outputFilePath);
                } catch (Exception e) {
                    //MessageBox.Show(e.Message);
                    eventsHandler.OnError(e.Message);
                }
            }
        }

        public void MathMLMultiple(ArrayList subList) {
            for (int i = 0; i < subList.Count; i++) {
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
        public String CheckFileOPen(ArrayList listSubDocs) {
            String resultSubDoc = "notopen";
            openSubdocs = new ArrayList();
            for (int i = 0; i < listSubDocs.Count; i++) {
                string[] splt = listSubDocs[i].ToString().Split('|');
                try {
                    Package pack;
                    pack = Package.Open(splt[0].ToString(), FileMode.Open, FileAccess.ReadWrite);
                    pack.Close();
                } catch {
                    resultSubDoc = "open";
                    openSubdocs.Add(splt[0].ToString() + "\n");
                }
            }
            return resultSubDoc;
        }

        #endregion

        #region Dynamic Menu
        public string GetDescriptionSingle(IRibbonControl control) {
            return "DAISY XML only (dtbook dtd)";
        }

        public string GetDescriptionSingleDTBook(IRibbonControl control) {
            return "Full text and full audio in MP3";
        }

        public string GetDescriptionMultiple(IRibbonControl control) {
            return "DAISY XML only (dtbook dtd)";
        }
        public string GetDescriptionMultipleDTBook(IRibbonControl control) {
            return "Full text and full audio in MP3";
        }

        /**
         * Dynamic conversion menus construction (called in the UI xml)
         * Note that the "_postprocess" script is excluded and reserved to dtbook post process when exporting the document
         */
        public string GetDTbook(IRibbonControl control) {
            StringBuilder MyStringBuilder = new StringBuilder(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");

            //PipelineConversionScripts = new Pipeline(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms\scripts");
            if (control.Id == "DaisyDTBookSingle" || control.Id == "DaisyTabDTBookSingle") {
                foreach (KeyValuePair<string, FileInfo> k in PipelineConversionScripts.ScriptsInfo) if (!k.Key.Equals("_postprocess"))  {
                    
                    PipelineMenuItem = new ToolStripMenuItem();
                    PipelineMenuItem.Text = k.Key;
                    PipelineMenuItem.AccessibleName = k.Key;
                    String quote = "<button id=\"" + k.Key.Replace(" ","_") + "\" tag=\"" + PipelineMenuItem.Text + "\" label=\"" + "&amp;" + PipelineMenuItem.Text + "\" onAction=\"" + "SaveAsDaisy" + "\"/>";
                    MyStringBuilder.Append(quote);
                }
                MyStringBuilder.Append(@"</menu>");
            } else if (control.Id == "DaisyDTBookMultiple" || control.Id == "DaisyTabDTBookMultiple") {
                foreach (KeyValuePair<string, FileInfo> k in PipelineConversionScripts.ScriptsInfo) if (!k.Key.Equals("_postprocess")) {
                    PipelineMenuItem = new ToolStripMenuItem();
                    PipelineMenuItem.Text = k.Key;
                    PipelineMenuItem.AccessibleName = k.Key;
                    String quote = "<button id=\"" + k.Key.Replace(" ", "_") + "\" tag=\"" + PipelineMenuItem.Text + "\" label=\"" + "&amp;" + PipelineMenuItem.Text + "\" onAction=\"" + "Mutiple" + "\"/>";
                    MyStringBuilder.Append(quote);
                }
                MyStringBuilder.Append(@"</menu>");
            }

            return MyStringBuilder.ToString();
        }
        #endregion

        #region ShapesObject

        public void saveasshapes() {
            MSword.Document doc = this.applicationObject.ActiveDocument;
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            List<string> warnings = new List<string>();
            String fileName = doc.Name.Replace(" ", "_");
            foreach (MSword.Shape item in doc.Shapes) {
                if (!item.Name.Contains("Text Box")) { 
                    object missing = Type.Missing;
                    item.Select(ref missing);
                    
                    string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-Shape" + item.ID.ToString() + ".png";
                    this.applicationObject.Selection.CopyAsPicture();
                    try {
                        // Note : using Clipboard.GetImage() set Word to display a clipboard data save request on closing
                        // So we rely on the user32 clipboard methods that does not seem to be intercepted by Office
                        System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                        byte[] Ret;
                        MemoryStream ms = new MemoryStream();
                        image.Save(ms, ImageFormat.Png);
                        Ret = ms.ToArray();
                        FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
                        fs.Write(Ret, 0, Ret.Length);
                        fs.Flush();
                        fs.Dispose();
                    } catch (ClipboardDataException cde) {
                        warnings.Add("- Shape " + item.ID + ": " + cde.Message);
                    } catch (Exception e) {
                        throw e;
                    } finally {
                        Clipboard.Clear();
                    }

                }
            }

            this.applicationObject.ActiveDocument.Save();

            if (warnings.Count > 0) {
                string warningMessage = "Some shapes could not be exported from the document:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
            
        }

        public void saveasshapes(ArrayList subList, String saveFlag) {
            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            List<string> warnings = new List<string>();
            for (int i = 0; i < subList.Count; i++) {
                object newName = subList[i];
                String fileName = newName.ToString().Replace(" ", "_");
                MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                foreach (MSword.Shape item in newDoc.Shapes) {
                    if (!item.Name.Contains("Text Box")) {
                        item.Select(ref missing);
                        string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-Shape" + item.ID.ToString() + ".png";
                        this.applicationObject.Selection.CopyAsPicture();
                        try {
                            System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                            byte[] Ret;
                            MemoryStream ms = new MemoryStream();
                            image.Save(ms, ImageFormat.Png);
                            Ret = ms.ToArray();
                            FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
                            fs.Write(Ret, 0, Ret.Length);
                            fs.Flush();
                            fs.Dispose();
                        } catch (ClipboardDataException cde) {
                            warnings.Add("- Shape " + item.ID.ToString() + ": " + cde.Message);
                        } catch (Exception e) {
                            throw e;
                        } finally {
                            Clipboard.Clear();
                            
                        }
                    }
                }
                
                newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            }
            if (warnings.Count > 0) {
                string warningMessage = "Some shapes could not be exported from the documents:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }
        /// <summary>
        /// Save the inline shapes 
        /// (Not e : not of type Embedded OLE or Pictures, so it is probably targeting inlined vectorial drawings)
        /// </summary>
        public void SaveasImages() {
            MSword.Document doc = this.applicationObject.ActiveDocument;

            object missing = Type.Missing;
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            List<string> warnings = new List<string>();
            String fileName = doc.Name.Replace(" ", "_");
            MSword.Range rng;
            foreach (MSword.Range tmprng in doc.StoryRanges) {
                rng = tmprng;
                while (rng != null) {
                    foreach (MSword.InlineShape item in rng.InlineShapes) {
                        if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture"))) {
                            object range = item.Range;
                            string str = "Shapes_" + GenerateId().ToString();
                            string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png";

                            item.Range.Bookmarks.Add(str, ref range);
                            item.Range.CopyAsPicture();
                            try {
                                System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                                byte[] Ret;
                                MemoryStream ms = new MemoryStream();
                                image.Save(ms, ImageFormat.Png);
                                Ret = ms.ToArray();
                                FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
                                fs.Write(Ret, 0, Ret.Length);
                                fs.Flush();
                                fs.Dispose();
                            } catch (ClipboardDataException cde) {
                                warnings.Add("- Image with AltText \"" + item.AlternativeText.ToString() + "\": " + cde.Message) ; 
                            } catch (Exception e) {
                                throw e;
                            } finally {
                                Clipboard.Clear();
                            }
                        }
                    }
                    rng = rng.NextStoryRange;
                }
            }
            
            this.applicationObject.ActiveDocument.Save();

            if (warnings.Count > 0) {
                string warningMessage = "Some images could not be exported from the document:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        public void SaveasImages(ArrayList subList, String saveFlag) {
            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();
            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
            List<string> warnings = new List<string>();
            for (int i = 0; i < subList.Count; i++) {
                object newName = subList[i].ToString();
                MSword.Document newDoc = this.applicationObject.Documents.Open(ref newName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                MSword.Range rng;
                String fileName = newName.ToString().Replace(" ", "_");
                foreach (MSword.Range tmprng in newDoc.StoryRanges) {
                    rng = tmprng;
                    while (rng != null) {
                        foreach (MSword.InlineShape item in rng.InlineShapes) {
                            if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture"))) {
                                string str = "Shapes_" + GenerateId().ToString();
                                string pathShape = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(fileName) + "-" + str + ".png";

                                object range = item.Range;
                                
                                item.Range.Bookmarks.Add(str, ref range);
                                item.Range.CopyAsPicture();

                                try {
                                    System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                                    byte[] Ret;
                                    MemoryStream ms = new MemoryStream();
                                    image.Save(ms, ImageFormat.Png);
                                    Ret = ms.ToArray();
                                    FileStream fs = new FileStream(pathShape, FileMode.Create, FileAccess.Write);
                                    fs.Write(Ret, 0, Ret.Length);
                                    fs.Flush();
                                    fs.Dispose();
                                } catch (ClipboardDataException cde) {
                                    warnings.Add("- Image with AltText \"" + item.AlternativeText.ToString() + "\": " + cde.Message);
                                } catch (Exception e) {
                                    throw e;
                                } finally {
                                    Clipboard.Clear();
                                    
                                }
                            }
                        }
                        rng = rng.NextStoryRange;
                    }
                }
                
                newDoc.Save();
                newDoc.Close(ref saveChanges, ref originalFormat, ref missing);
            }

            if (warnings.Count > 0) {
                string warningMessage = "Some images could not be exported from the documents:";
                foreach (string warning in warnings) {
                    warningMessage += "\r\n" + warning;
                }
                throw new Exception(warningMessage);
            }
        }

        #endregion

        #region FootNote

        public void AddFotenote(IRibbonControl control) {
            try {
                MSword.Document docActive = this.applicationObject.ActiveDocument;

                Microsoft.Office.Interop.Word.Range footnotetext = docActive.Application.Selection.Range;

                if (footnotetext.Text == null) {
                    MessageBox.Show("Please select Footnote text from the document", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } else {
                    String selectedText = docActive.Application.Selection.Text;
                    String refNumber = string.Empty;
                    String ftNtText = string.Empty;
                    Object refNumberObj = null;

                    //To consider any text provided as index of reference
                    string[] footNote = selectedText.Split(' ');
                    if (footNote.Length <= 1) {
                        throw new Exception("Selected Text is not a valid Footnote");
                    }
                    refNumber = footNote[0];
                    //Remove dot at the end, if any
                    refNumber = refNumber.TrimEnd('.');
                    refNumberObj = refNumber;

                    ftNtText = string.Empty;
                    for (int i = 1; i < footNote.Length; i++) {
                        ftNtText = ftNtText + footNote[i] + " ";
                    }
                    Object actualText = ftNtText;
                    ftNtText = actualText.ToString();

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
                    //fndValue.MatchSuffix = true;
                    //Shaby (end)

                    footntRefrns = new ArrayList();
                    ExecuteFind(fndValue);
                    int occurance = 0;
                    while (fndValue.Found) {
                        int Strng = rngDoc.Start;
                        Object startRng;
                        startRng = Strng - 30;
                        if (Convert.ToInt32(startRng) <= 0) {
                            startRng = 0;
                        }
                        int rdRng = rngDoc.End;
                        Object endRng = rdRng + 20;
                        if (Convert.ToInt32(endRng) >= Convert.ToInt32(endValue)) {
                            endRng = endValue;
                        }
                        MSword.Range getRng = docActive.Range(ref startRng, ref endRng);

                        if (getRng.Text != null) {
                            //Shaby : Below check done to avoid selecting the actual footer
                            if (!rngDoc.InRange(footnotetext)) {
                                //Shaby : select the apt index, if there are multiple Footnote index in selected body text 
                                if (getRng.Text.ToLower().IndexOf(fndValue.Text) != getRng.Text.ToLower().LastIndexOf(fndValue.Text)) {
                                    //Trimming the left position if any index exist in the left part
                                    int leftDistanceToRef = rngDoc.Start - getRng.Start;
                                    string leftTextToRef = getRng.Text.Substring(0, leftDistanceToRef);
                                    if (leftTextToRef.ToLower().Contains(fndValue.Text.ToLower())) {
                                        startRng = rngDoc.Start;
                                    }
                                    //Trimming the right position if any index exists in the right part
                                    int rightDistanceFromRef = getRng.End - rngDoc.End;
                                    int leftDistanceFromRef = rngDoc.End - getRng.Start;
                                    string rightTextFromRef = getRng.Text.Substring(leftDistanceFromRef, rightDistanceFromRef);
                                    if (rightTextFromRef.ToLower().Contains(fndValue.Text.ToLower())) {
                                        endRng = rngDoc.End;
                                    }
                                }


                                ////Shaby2: Trim till the last linebreak(if any), at both the sides
                                MSword.Range lineBreakerRng = docActive.Range(ref startRng, ref endRng);
                                while ((lineBreakerRng.Text.IndexOf("\r")) != -1) {
                                    int indexOfFooterIndex = lineBreakerRng.Text.ToLower().IndexOf(fndValue.Text);
                                    if ((lineBreakerRng.Text.IndexOf("\r")) < indexOfFooterIndex) {
                                        startRng = lineBreakerRng.Start + lineBreakerRng.Text.IndexOf("\r") + 1;
                                    } else {
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
                    if (footntRefrns.Count != 0) {
                        SuggestedReferences suggestForm = new SuggestedReferences(docActive, footntRefrns, ftNtText);
                        suggestForm.ShowDialog();
                    } else {
                        MessageBox.Show("No match found in the document to map selected footnote", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            } catch (Exception x) {
                MessageBox.Show(x.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public Boolean ExecuteFind(MSword.Find find) {
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
        private string ctrlId;

        

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

        private void MTGetEquationAddin(IPluginEventsHandler eventsHandler) {
            Microsoft.Office.Interop.Word.Document doc = this.applicationObject.ActiveDocument;

            if (doc != null) {
                int iNumShapesIterated = 0;
                iNumShapesIterated = IterateShapes(eventsHandler);
            }
        }

        private int IterateShapes(IPluginEventsHandler eventsHandler)//Hashtable multipleMathMl, IPluginEventsHandler eventsHandler)
        {
            Int16 showMsg = 0;
            MSword.Range rng;
            String storyName = "";
            int iNumShapesViewed = 0;
            Microsoft.Office.Interop.Word.Document doc = this.applicationObject.ActiveDocument;
            //Make sure that we have shapes to iterate over

            foreach (MSword.Range tmprng in doc.StoryRanges) {
                listmathML = new ArrayList();
                rng = tmprng;
                storyName = rng.StoryType.ToString();
                while (rng != null) {
                    storyName = rng.StoryType.ToString();
                    MSword.InlineShapes shapes = rng.InlineShapes;
                    if (shapes != null && shapes.Count > 0) {
                        int iCount = 1;
                        int iNumShapes = 0;
                        Microsoft.Office.Interop.Word.InlineShape shape;
                        iNumShapes = shapes.Count;
                        //iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.

                        while (iCount <= iNumShapes) {
                            if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject") {
                                if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4") {
                                    shape = shapes[iCount];

                                    if (shape != null && shape.OLEFormat != null) {
                                        bool bRetVal = false;
                                        string strProgID;
                                        Guid autoConvert;
                                        strProgID = shape.OLEFormat.ProgID;
                                        bRetVal = FindAutoConvert(ref strProgID, out autoConvert);

                                        // if we are successful with the conversion of the CLSID we now need to query
                                        //  the application to see if it can actually do the work
                                        if (bRetVal == true) {
                                            bool bInsertable = false;
                                            bool bNotInsertable = false;

                                            bInsertable = IsCLSIDInsertable(ref autoConvert);
                                            bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

                                            //Make sure that the server of interest is insertable and not-insertable
                                            if (bInsertable && bNotInsertable) {
                                                bool bServerExists = false;
                                                string strPathToExe;
                                                bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

                                                //if the server exists then see if MathML can be retrieved for the shape
                                                if (bServerExists) {
                                                    bool bMathML = false;
                                                    string strVerb;
                                                    int indexForVerb = -100;

                                                    strVerb = "RunForConversion";

                                                    bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
                                                    if (bMathML) {
                                                        Equation_GetMathML(ref shape, indexForVerb);
                                                    }
                                                }
                                            } else {
                                                if (bInsertable != bNotInsertable) {
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
            if (showMsg == 1) {
                string message =
                    "In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images";
                eventsHandler.OnStop(message, "Warning");
            }
            //System.Windows.Forms.MessageBox.Show("In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

            return iNumShapesViewed;
        }

        private bool FindAutoConvert(ref string ProgID, out Guid autoConvert) {
            bool bRetVal = false;
            Guid oGuid;
            int iCOMRetVal = 0;

            iCOMRetVal = CLSIDFromProgID(ProgID, out oGuid);

            if (iCOMRetVal == 0) {
                RecurseAutoConvert(ref oGuid, out autoConvert);
                bRetVal = true;
            } else {
                autoConvert = oGuid;
            }

            return bRetVal;
        }

        private void RecurseAutoConvert(ref Guid oGuid, out Guid autoConvert) {
            int iCOMRetVal = 0;

            iCOMRetVal = OleGetAutoConvert(ref oGuid, out autoConvert);
            if (iCOMRetVal == 0) {
                //If we have no error and the the CLSIDs are the same, then make sure that
                // 
                bool bGuidTheSame = false;
                try {
                    bGuidTheSame = IsEqualGUID(ref oGuid, ref autoConvert);
                } catch (COMException eCOM) {
                    MessageBox.Show(eCOM.Message);
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                }

                if (bGuidTheSame == false) {
                    oGuid = autoConvert;
                    RecurseAutoConvert(ref oGuid, out autoConvert);
                }
            } else {
                //There was some error in the auto conversion.
                // See if this guid will do the conversion.
                autoConvert = oGuid;
            }
        }

        private bool IsCLSIDInsertable(ref Guid oGuid) {
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

        private bool IsCLSIDNotInsertable(ref Guid oGuid) {
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

        private bool DoesServerExist(out string strPathToExe, ref Guid oGuid) {
            bool bServerExists = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"LocalServer32";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetValueNames();
                strPathToExe = "";
                try {
                    strPathToExe = (string)regkey.GetValue(valnames[0]);
                } catch (Exception e) {
                }

                if (strPathToExe.Length > 0) {
                    //Now check if this is a good path
                    if (File.Exists(strPathToExe))
                        bServerExists = true;
                }

            } else {

                strPathToExe = null;

            }

            return bServerExists;
        }

        private bool DoesServerSupportMathML(ref Guid oGuid, ref string strVerb, out int indexForVerb) {
            bool bIsMathMLSupported = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\DataFormats\GetSet";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetSubKeyNames();
                int x = 0;
                while (x < regkey.SubKeyCount) {
                    RegistryKey subKey;
                    if (regkey.SubKeyCount > 0) {
                        subKey = regkey.OpenSubKey(valnames[x]);
                        if (subKey != null) {
                            string[] dataFormats = subKey.GetValueNames();
                            int y = 0;
                            while (y < subKey.ValueCount) {
                                string strValue = (string)subKey.GetValue(dataFormats[y]);

                                //This will accept both MathML and MathML Presentation.
                                if (strValue.Contains("MathML")) {
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
            if (bIsMathMLSupported) {
                //The return value for a verb not found will be 1000
                //
                indexForVerb = GetVerbIndex(strVerb, ref oGuid);

                if (indexForVerb == 1000) {
                    bIsMathMLSupported = false;
                }
            } else {
                //We do not have an appropriate verb to start the server
                indexForVerb = -100;  //There is a predefined range for 
            }

            return bIsMathMLSupported;
        }

        private int GetVerbIndex(string strVerbToFind, ref Guid oGuid) {
            int indexForVerb = 1000;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\Verb";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                //Lets make sure that we have some values before preceeding.
                if (regkey.SubKeyCount > 0) {
                    int x = 0;
                    int iCount = 0;

                    string[] valnames = regkey.GetSubKeyNames();

                    while (x < regkey.SubKeyCount) {
                        RegistryKey subKey;
                        if (regkey.SubKeyCount > 0) {
                            subKey = regkey.OpenSubKey(valnames[x]);
                            if (subKey != null) {
                                int y = 0;
                                string[] verbs = subKey.GetValueNames();
                                iCount = subKey.ValueCount;
                                string verb;

                                //Search all of the verbs for requested string.
                                while (y < iCount) {
                                    verb = (string)subKey.GetValue(verbs[y]);
                                    if (verb.Contains(strVerbToFind) == true) {
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

        private void Equation_GetMathML(ref Microsoft.Office.Interop.Word.InlineShape shape, int indexForVerb) {
            if (shape != null) {
                object dataObject = null;
                object objVerb;

                objVerb = indexForVerb;

                //Start MathType, and get the dataobject that is connected to the server.    
                shape.OLEFormat.DoVerb(ref objVerb);

                try {
                    dataObject = shape.OLEFormat.Object;
                } catch (Exception e) {
                    //we have an issue with trying to get the verb,
                    //  There will be a attempt at another way to start the application.
                    MessageBox.Show(e.Message);
                }

                IOleObject oleObject = null;

                //This is a C# version of a QueryInterface
                if (dataObject != null) {
                    mDataObject = dataObject as IConnectDataObject;


                    oleObject = dataObject as IOleObject;
                } else {
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

                if (mDataObject != null) {
                    int iRetVal = 0;

                    //Initialize a FORMATETC structure to get the requested data
                    oFormatEtc.cfFormat = (Int16)oFormat.Id;
                    oFormatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
                    oFormatEtc.lindex = -1;
                    oFormatEtc.ptd = (IntPtr)0;
                    oFormatEtc.tymed = TYMED.TYMED_HGLOBAL;

                    iRetVal = mDataObject.QueryGetData(ref oFormatEtc);
                    //iRetVal will be zero if the MathML type is contained within the server.
                    if (iRetVal == 0) {
                        oStgMedium.tymed = TYMED.TYMED_NULL;
                    }

                    try {
                        mDataObject.GetData(ref oFormatEtc, out oStgMedium);
                    } catch (System.Runtime.InteropServices.COMException e) {
                        System.Windows.Forms.MessageBox.Show(e.ToString());
                        throw;
                    }

                    // Because we explicitly requested a MathML, we know that it is TYMED_HGLOBAL
                    // lets deal with the memory here.
                    if (oStgMedium.tymed == TYMED.TYMED_HGLOBAL &&
                        oStgMedium.unionmember != null) {
                        WriteOutMathMLFromStgMedium(ref oStgMedium);

                        if (oleObject != null) {
                            uint close = (uint)OLECLOSE.OLECLOSE_NOSAVE;
                            // uint close = (uint)Microsoft.VisualStudio.OLE.Interop.OLECLOSE.OLECLOSE_NOSAVE;
                            oleObject.Close(close);
                        }
                    }
                }
            }
        }

        private void WriteOutMathMLFromStgMedium(ref ConnectSTGMEDIUM oStgMedium) {
            IntPtr ptr;
            byte[] rawArray = null;


            //Verify that our data contained within the STGMEDIUM is non-null
            if (oStgMedium.unionmember != null) {
                //Get the pointer to the data that is contained
                //  within the STGMEDIUM
                ptr = oStgMedium.unionmember;

                //The pointer now becomes a Handle reference.
                HandleRef handleRef = new HandleRef(null, ptr);

                try {
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
                } catch (Exception exp) {
                    System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + exp.ToString());
                } finally {
                    //This gets called regardless within a try catch.
                    //  It is a good place to clean up like this.
                    GlobalUnlock(handleRef);
                }
            }
        }

        private void MTGetEquationAddinNew(String newName) {
            Object fileName = newName;
            object addToRecentFiles = false;
            object readOnly = false;
            object isVisible = false;
            object missing = Type.Missing;

            try {
                Microsoft.Office.Interop.Word.Document doc = this.applicationObject.Documents.Open(ref fileName, ref missing, ref readOnly, ref addToRecentFiles, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);

                if (doc != null) {
                    int iNumShapesIterated = 0;
                    iNumShapesIterated = IterateShapesNew(doc);
                }
            } catch (Exception e) {

            }
        }

        private int IterateShapesNew(MSword.Document doc) {
            Int16 showMsg = 0;
            int iNumShapesViewed = 0;
            MSword.Range rng;
            String storyName = "";
            object missing = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;

            foreach (MSword.Range tmprng in doc.StoryRanges) {
                listmathML = new ArrayList();
                rng = tmprng;
                storyName = rng.StoryType.ToString();
                while (rng != null) {
                    MSword.InlineShapes shapes = rng.InlineShapes;
                    if (shapes != null && shapes.Count > 0) {
                        int iCount = 1;
                        int iNumShapes = 0;
                        Microsoft.Office.Interop.Word.InlineShape shape;
                        iNumShapes = shapes.Count;
                        //iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.
                        while (iCount <= iNumShapes) {
                            if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject") {
                                if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4") {
                                    shape = shapes[iCount];

                                    if (shape != null && shape.OLEFormat != null) {
                                        bool bRetVal = false;
                                        string strProgID;
                                        Guid autoConvert;
                                        strProgID = shape.OLEFormat.ProgID;

                                        bRetVal = FindAutoConvert(ref strProgID, out autoConvert);
                                        // if we are successful with the conversion of the CLSID we now need to query
                                        //  the application to see if it can actually do the work
                                        if (bRetVal == true) {
                                            bool bInsertable = false;
                                            bool bNotInsertable = false;

                                            bInsertable = IsCLSIDInsertable(ref autoConvert);
                                            bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

                                            //Make sure that the server of interest is insertable and not-insertable
                                            if (bInsertable && bNotInsertable) {
                                                bool bServerExists = false;
                                                string strPathToExe;
                                                bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

                                                //if the server exists then see if MathML can be retrieved for the shape
                                                if (bServerExists) {
                                                    bool bMathML = false;
                                                    string strVerb;
                                                    int indexForVerb = -100;

                                                    strVerb = "RunForConversion";

                                                    bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
                                                    if (bMathML) {
                                                        Equation_GetMathML(ref shape, indexForVerb);
                                                    }
                                                }
                                            } else {
                                                if (bInsertable != bNotInsertable) {
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

        public string MasterSubDecision(string tempInput, IPluginEventsHandler eventsHandler) {
            string masterSubFlag = string.Empty;
            Package pack = Package.Open(tempInput, FileMode.Open, FileAccess.ReadWrite);

            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                packRelationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);
            int cnt = 0;
            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships()) {
                packRelationship = searchRelation;
                if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument") {
                    if (packRelationship.TargetMode.ToString() == "External") {
                        cnt++;
                    }
                }
            }
            pack.Close();

            if (cnt > 0) {
                masterSubFlag = eventsHandler.AskForTranslatingSubdocuments() ? "Yes" : "No";
            } else
                masterSubFlag = "NoMasterSub";

            return masterSubFlag;
        }

        /// <summary>
        /// Prepare conversion of subdocuments
        /// </summary>
        /// <param name="preparetionResult"></param>
        /// <param name="cTrl"></param>
        /// <param name="eventsHandler"></param>
        /// <returns></returns>
        public bool OoxToDaisyOwn(PreparetionResult preparetionResult, String cTrl, IPluginEventsHandler eventsHandler) {
            SubdocumentsList subdocuments = SubdocumentsManager.FindSubdocuments(preparetionResult.TempFilePath, preparetionResult.OriginalFilePath);
            //MessageBox.Show("Check for errors when retrieving subdocuments pathes");
            if(subdocuments.Errors.Count > 0) {
                StringBuilder errorMessage = new StringBuilder();
                errorMessage.Append("Errors were encoutered while retrieving sub documents:");
                foreach (string error in subdocuments.Errors) {
                    errorMessage.Append("\r\n- " + error);
                }
                eventsHandler.OnError(errorMessage.ToString());
                return false;
            }

            //MessageBox.Show("Get not translated docs");
            notTranslatedDoc = subdocuments.GetNotTraslatedSubdocumentsNames();

            //MessageBox.Show("Doc sublist");
            ArrayList subList = new ArrayList();
            subList.Add(preparetionResult.TempFilePath + "|Master");
            foreach (string subdoc in subdocuments.GetSubdocumentsNameWithRelationship()) {
                subList.Add(subdoc);
            }
            
            int subCount = subdocuments.SubdocumentsCount + 1;
            //Checking whether any original or Subdocumets is already Open or not
            string resultOpenSub = CheckFileOPen(subList);
            if (resultOpenSub != "notopen") {
                eventsHandler.OnError("Some Sub documents are in open state.\r\nPlease close all the Sub documents before Translation.");
                return false;
            }
            //Checking whether Sub documents are Simple documents or a Master document
            string resultSub = SubdocumentsManager.CheckingSubDocs(subdocuments.GetSubdocumentsNameWithRelationship());
            if (resultSub != "simple") {
                eventsHandler.OnError("Some of the added documents are MasterSub documents.Please add simple documents.");
                return false;
            }

            try {
                saveasshapes(subdocuments.GetSubdocumentsNames(), "Yes");
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }
            try {
                SaveasImages(subdocuments.GetSubdocumentsNames(), "Yes");
            } catch (Exception e) {
                eventsHandler.OnError("An error occured while preprocessing images and may prevent the rest of the conversion to success:\r\n" + e.Message);
            }

            MathMLMultiple(subList);
            this.applicationObject.ActiveDocument.Save();
            return true;
            //OoxToDaisyUI(preparetionResult, cTrl);
        }

        #endregion

        #region Language

        public void GetlanguageSettings(IRibbonControl control) {
            try {
                int fileIndex;
                MSword.Document currentDoc = this.applicationObject.ActiveDocument;
                if (!currentDoc.Saved || currentDoc.FullName.LastIndexOf('.') < 0) {
                    System.Windows.Forms.MessageBox.Show(addinLib.GetString("DaisySaveDocumentBeforelanguage"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                } else if (currentDoc.Saved) {
                    fileIndex = currentDoc.FullName.LastIndexOf('.');
                    String substr = currentDoc.FullName.Substring(fileIndex);

                    if (substr.ToLower() != ".docx") {
                        MessageBox.Show(addinLib.GetString("DaisySaveDocumentin2007"), "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
                    } else {
                        if (currentDoc.Application.Selection.Start != currentDoc.Application.Selection.End) {
                            ArrayList list = new ArrayList();
                            foreach (string item in Enum.GetNames(typeof(MSword.WdLanguageID))) {
                                MSword.WdLanguageID value = (MSword.WdLanguageID)Enum.Parse(typeof(MSword.WdLanguageID), item);
                                list.Add(value);
                            }

                            Language lng = new Language(list, currentDoc.Application.Selection.Range, currentDoc);
                            lng.ShowDialog();
                        } else {
                            MessageBox.Show("Select a paragraph for setting the Language", "SaveAsDAISY", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        }
                    }
                }
            } catch (Exception x) {
                MessageBox.Show(x.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion


    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleObject {
        void DoNotCall_1(object x);
        void DoNotCall_2(ref object x);
        void SetHostNames(object szContainerApp, object szContainerObj);
        void Close(uint dwSaveOption);
    };

    public enum OLECLOSE {
        OLECLOSE_SAVEIFDIRTY = 0,
        OLECLOSE_NOSAVE = 1,
        OLECLOSE_PROMPTSAVE = 2,
    }
}