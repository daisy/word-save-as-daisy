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
using Daisy.SaveAsDAISY;
using Daisy.SaveAsDAISY.Conversion;
using System.Globalization;
using Daisy.SaveAsDAISY.Conversion.Events;
using System.Text.RegularExpressions;
using Daisy.SaveAsDAISY.Forms;
using Script = Daisy.SaveAsDAISY.Conversion.Script;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts;
using System.Diagnostics;
using Microsoft.Toolkit.Uwp.Notifications;
using Daisy.SaveAsDAISY.WPF;


namespace Daisy.SaveAsDAISY.Addins.Word2007 {
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

        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        PackageRelationship packRelationship = null;
        XmlDocument currentDocXml, validation_xml;
        string docFile, path_For_Xml;
        public ArrayList docValidation = new ArrayList();
        private IRibbonUI daisyRibbon;
        private bool showValidateTabBool = false;
        frmValidate2007 validateInput;

        private MSword.Application applicationObject;
        private AddinResources addinLib;

        private ArrayList footntRefrns;


        /// <summary>
        ///	Implements the constructor for the Add-in object.
        ///	Place your initialization code within this method.
        /// </summary>
        public Connect() {
            try {
                this.addinLib = new AddinResources();
            } catch (Exception e) {
                AddinLogger.Error(e);
                MessageBox.Show(e.Message);
                //throw;
            }
            
        }

        private static readonly string NOTIFICATIONFILEPATH = Path.Combine(ConverterSettings.ApplicationDataFolder, "notify");
        private static readonly string SPONSORSHIPURL = "https://inclusivepublishing.org/sponsorship/";

        public void NotifyDonationRequest()
        {
            if (ConverterSettings.Instance.DontNotifySponsorship) return;

            int remainingConversionBeforeNotify = 0; // notify
            if (File.Exists(NOTIFICATIONFILEPATH)) {
                string previousConversionCounter = File.ReadAllText(NOTIFICATIONFILEPATH);
                int.TryParse(previousConversionCounter, out remainingConversionBeforeNotify);
            }
            if (remainingConversionBeforeNotify > 0) {
                remainingConversionBeforeNotify--;
            } else {
                string logoPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "daisy_high.ico");
                try {
                    // Notification for donations request to support the developement
                    new ToastContentBuilder()
                        .AddAttributionText("The DAISY Consortium")
                        .AddAppLogoOverride(new Uri(logoPath))
                        .AddText("If you find SaveAsDAISY useful, please help us by donating to support its ongoing maintenance.")
                        .AddButton("Support our work", ToastActivationType.Background, SPONSORSHIPURL)
                        .SetProtocolActivation(new Uri(SPONSORSHIPURL))
                        .Show();
                } catch (Exception e) {
                    AddinLogger.Error(e);
                }
               
               
                remainingConversionBeforeNotify = 10;
            }
            File.WriteAllText(NOTIFICATIONFILEPATH, remainingConversionBeforeNotify.ToString());
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
                
                // Listen to notification activation
                ToastNotificationManagerCompat.OnActivated += toastArgs =>
                {
                    File.WriteAllText(NOTIFICATIONFILEPATH, 50.ToString());
                    Process.Start(SPONSORSHIPURL);
                };
            } catch (Exception e) {
                AddinLogger.Error(e);
                MessageBox.Show(e.Message);
                //throw;
            }
        }

        /// <summary>
        /// Only activate the action controls on non-protected view
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public bool ribbonButtonsEnabled(IRibbonControl control) {
            if(this.applicationObject.ActiveProtectedViewWindow != null) {
                return false;
            }
            return true;
        }
        
        /// <summary>
        /// Function to do changes in Look and Feel of UI
        /// </summary>
        void applicationObject_DocumentChange() {
            
            if(this.applicationObject.Documents.Count > 0) {
                CheckforAttchments(this.applicationObject.ActiveDocument);
            }
            showValidateTabBool = false;
            if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");

        }
        //Envent handling deactivation of word document
        void applicationObject_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn) {

        }
        /// <summary>
        ///Core Function to check the Daisy Styles in the current Document
        /// </summary>
        /// <param name="Doc">Current Document</param>
        void applicationObject_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc) {
            CheckforAttchments(Doc);
        }


        /// <summary>
        /// Function to check whether the Daisy Styles are new or not
        /// </summary>
        public void CheckforAttchments(MSword.Document doc) {
            Dictionary<int, string> objArray = new Dictionary<int, string>();
            for (int iCountStyles = 1; iCountStyles <= doc.Styles.Count; iCountStyles++) {
                object ActualVal = iCountStyles;
                string strValue = doc.Styles.get_Item(ref ActualVal).NameLocal.ToString();
                if (strValue.EndsWith("(DAISY)")) {
                    objArray.Add(iCountStyles, strValue);
                }
            }
            showViewTabBool = objArray.Count == AddInHelper.DaisyStylesCount;
            if (daisyRibbon != null) daisyRibbon.InvalidateControl("ImportDaisyStylesTabButton");
        }

        /// <summary>
        /// Function to remove Unwanted bookmarks before saving the document
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

            // Discard notifications from the addin when word/the addin is closing
            // (avoid a relaunch of word and the addin to handle the notification action)
            ToastNotificationManagerCompat.Uninstall();
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

        /// <summary>
        /// Select the UI XML description of the Ribbon, 
        /// depending on word version and if a pipeline distribution is present near the assembly
        /// </summary>
        /// <param name="RibbonID"></param>
        /// <returns></returns>
        string IRibbonExtensibility.GetCustomUI(string RibbonID) {
            try
            {
                if (Directory.Exists(ConverterHelper.Pipeline2Path))
                {
                    // Removing the validator button for office 2010 and later
                    // For office 2007, the old validation step remains necessary
                    if (this.applicationObject.Version == "12.0")
                    {
                        return GetResource("customUI2007.xml");
                    }
                    else return GetResource("customUI.xml");
                }
                else
                {
                    return GetResource("customUIOld.xml");
                }
            } catch (Exception e)
            {
                AddinLogger.Error(e);
            }
            return null;
        }

        

        /// <summary>
        /// Retrieve the content of a textual resource from the assembly manifest
        /// (Note: Only used to retrieve UI XML)
        /// </summary>
        /// <param name="resourceName"></param>
        /// <returns></returns>
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

        #region Icons
        /// <summary>
        /// Icons filenames mapping with ribbon control IDs
        /// </summary>
        private static readonly Dictionary<string, string> controlsIconNames = new Dictionary<string, string> () {
            { "Daisy", "speaker.jpg" },
            { "SaveAsDaisyOfficeMenu", "speaker.jpg" },
            { "ExportToXMLOfficeButton", "Singlexml.png" },
            { "ExportToDaisy3OfficeButton", "Singlexml.png" },
            { "ExportToEpub3OfficeButton", "speaker.jpg" },
            { "ExportToMP3OfficeButton", "speaker.jpg" },
            { "ExportToDaisy202OfficeButton", "speaker.jpg" },
            { "DaisyMultiple", "Multiplexml.png" },
            { "DaisyDTBookMultiple", "subfolder.png" },
            { "ExportTabMenu", "speaker.jpg" },
            { "ExportToXMLTabButton", "Singlexml.png" },
            { "ExportToDaisy3TabButton", "speaker.jpg" },
            { "ExportToEpub3TabButton", "speaker.jpg" },
            { "ExportToMP3TabButton", "speaker.jpg" },
            { "ExportToDaisy202TabButton", "speaker.jpg" },
            { "DtbookTabMultiple", "Multiplexml.png" },
            { "DaisyTabMultiple", "Multiplexml.png" },
            { "EpubTabMultiple", "subfolder.png" },
            { "Button1", "speaker.jpg" },
            { "Button2", "subfolder.png" },
            { "MarkAbbreviationTabButton", "ABBR.png" },
            { "ManageAbbreviationsTabButton", "ABBR2.png" },
            { "MarkAcronymTabButton", "ACR.png" },
            { "ManageAcronymsTabButton", "ACR2.png" },
            { "toggleValidateTabButton", "validate.png" },
            { "ImportDaisyStylesTabButton", "import.png" },
            { "AddFootnotesTabButton", "footnotes.png" },
            { "DocumentLanguageTabButton", "Language.png" },
            { "SettingsTabButton", "gear.png" },
            { "VersionDetailsTabButton", "version.png" },
            { "OpenManualTabButton", "help.png" },
            { "DocumentationTabMenu", "speaker.jpg" }
        };
        public stdole.IPictureDisp iconSelector(IRibbonControl control) {
            try
            {
                return this.addinLib.GetLogo(controlsIconNames[control.Id]);
            } catch (Exception e)
            {
                AddinLogger.Error(e);
                throw e;
            }
            
        }

        #endregion Icons

        #region Labels and Descriptions
        public string getRibbonLabel(IRibbonControl control)
        {
            try
            {
                return RibbonLabels.ResourceManager.GetString(control.Id);
            }
            catch (Exception e)
            {
                AddinLogger.Error(e);
                throw e;
            }

        }

        public string getRibbonDescription(IRibbonControl control)
        {
            try
            {
                return RibbonDescriptions.ResourceManager.GetString(control.Id);
            }
            catch (Exception e)
            {
                AddinLogger.Error(e.Message);
                throw e;
            }

        }

        #endregion Labels and Descriptions

        public void OnLoad(IRibbonUI ribbon) {
            daisyRibbon = ribbon;
        }
        private bool showViewTabBool = false;


        public void GetDaisySettings(IRibbonControl control) {
            //ConverterSettingsForm daisyfrm = new ConverterSettingsForm();
            //daisyfrm.ShowDialog();
            Daisy.SaveAsDAISY.WPF.Settings settings = new Daisy.SaveAsDAISY.WPF.Settings();
            settings.ShowDialog();
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
                docTemplate = null;
                this.applicationObject.NormalTemplate.Save();

                if (File.Exists(copyTempName.ToString())) {
                    File.Delete(copyTempName.ToString());
                }

                showViewTabBool = true;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("ImportDaisyStylesTabButton");

            } catch (Exception ex) {
                string stre = ex.Message;
                MessageBox.Show(ex.Message.ToString(), "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                if(docTemplate != null)
                {
                    object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
                    docTemplate.Close(ref saveChanges, ref missing, ref missing);
                }
            }

        }


        public bool getEnabled(IRibbonControl control) {
            return this.ribbonButtonsEnabled(control) && !showViewTabBool;
        }
        //Method for Disabling validate tab
        public bool buttonValidatePressed(IRibbonControl control) {
            return showValidateTabBool;
        }
        //Method for Enabling validate tab
        public bool getValidateEnabled(IRibbonControl control) {
            return !showValidateTabBool;
        }

        /*Function validate input docx file using DAISY rule*/
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
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
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
                    if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
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
                        if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
                        return;
                    }
                    //if docx format
                    else {
                        docFile = this.addinLib.GetTempPath((string)doc.FullName, ".docx");
                        File.Copy((string)doc.FullName, (string)docFile);
                    }
                    DeleteBookMark(this.applicationObject.ActiveDocument);

                    object saveChangesWord = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                    object originalFormatWord = Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat;
                    object FileName = (object)path;
                    object AddToRecentFiles = false;
                    object routeDocument = Type.Missing;
                    //object Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
                    //Getting validation xml from executing assembly
                    path_For_Xml = ConverterHelper.AppDataSaveAsDAISYDirectory;

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
                                    if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
                                }
                            }
                            //Progress bar canceled
                            if (dr == System.Windows.Forms.DialogResult.Cancel) {
                                MessageBox.Show("Validation stopped", "Quit validation", System.Windows.Forms.MessageBoxButtons.OK,
                                                System.Windows.Forms.MessageBoxIcon.Stop);
                                showValidateTabBool = false;
                                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
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
                DeleteBookMark(this.applicationObject.ActiveDocument);
                showValidateTabBool = false;
                if (daisyRibbon != null) daisyRibbon.InvalidateControl("toggleValidateTabButton");
                this.applicationObject.ActiveDocument.Save();
            } catch {

            }
        }


        /*Function which deletes bookmark */
        public void DeleteBookMark(Document doc) {
            try {
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
                AddinLogger.Error(ex);
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
            System.Diagnostics.Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "SaveAsDAISY_Instruction_Manual_July_2023.docx");
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
            Daisy.SaveAsDAISY.WPF.About aboutWindow = new Daisy.SaveAsDAISY.WPF.About();
            aboutWindow.ShowDialog();
            //About abtForm = new About();
            //abtForm.ShowDialog();
        }

        #region Single document conversion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pipelineScript"></param>
        /// <param name="eventsHandler"></param>
        /// <param name="conversionIntegrationTestSettings"></param>
        //public void ApplyScriptOld(Script pipelineScript, IConversionEventsHandler eventsHandler, ConversionParameters conversionIntegrationTestSettings = null)
        //{
        //    try
        //    {
        //        IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
        //        WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
        //        Converter converter = null;
        //        if (conversionIntegrationTestSettings != null)
        //        {
        //            ConversionParameters conversion = conversionIntegrationTestSettings.withParameter("Version", this.applicationObject.Version);
        //            converter = new Converter(preprocess, documentConverter, conversion, eventsHandler);
        //        }
        //        else
        //        {
        //            ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
        //            converter = new GraphicalConverter(preprocess, documentConverter, conversion, (GraphicalEventsHandler)eventsHandler);
        //        }

        //        DocumentParameters currentDocument = converter.PreprocessDocument(this.applicationObject.ActiveDocument.FullName);
        //        if (conversionIntegrationTestSettings != null
        //                || ((GraphicalConverter)converter).requestUserParameters(currentDocument) == ConversionStatus.ReadyForConversion)
        //        {
                    
        //            ConversionResult result = converter.Convert(currentDocument);
        //            // Conversion test with 
        //            //ConversionResult result = converter.Convert2(currentDocument);
        //            if (result != null && result.Succeeded)
        //            {
        //                Process.Start(
        //                    Directory.Exists(converter.ConversionParameters.OutputPath)
        //                    ? converter.ConversionParameters.OutputPath
        //                    : Path.GetDirectoryName(converter.ConversionParameters.OutputPath)
        //                );
        //            } else {
        //                MessageBox.Show(result.UnknownErrorMessage, "Conversion failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //            }

        //        }
        //        else
        //        {
        //            eventsHandler.onConversionCanceled();
        //        }

        //        //applicationObject.ActiveDocument.Save();
        //    }
        //    catch (Exception e)
        //    {
        //        AddinLogger.Error(e);
        //        if (conversionIntegrationTestSettings == null)
        //        {
        //            ExceptionReport report = new ExceptionReport(e);
        //            report.ShowDialog();
        //        }

        //    }
        //}

        public void ApplyScript(Script pipelineScript, IConversionEventsHandler eventsHandler, ConversionParameters conversionIntegrationTestSettings = null)
        {
            try {
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                Converter converter = null;
                if (conversionIntegrationTestSettings != null) {
                    ConversionParameters conversion = conversionIntegrationTestSettings.withParameter("Version", this.applicationObject.Version);
                    converter = new Converter(preprocess, documentConverter, conversion, eventsHandler);
                } else {
                    ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                    converter = new GraphicalConverter(preprocess, documentConverter, conversion, (GraphicalEventsHandler)eventsHandler);
                }

                DocumentParameters currentDocument = converter.PreprocessDocument(this.applicationObject.ActiveDocument.FullName);
                if (conversionIntegrationTestSettings != null
                        || ((GraphicalConverter)converter).requestUserParameters(currentDocument) == ConversionStatus.ReadyForConversion) {

                    //ConversionResult result = converter.Convert(currentDocument);
                    // Conversion test with 
                    ConversionResult result = converter.ConvertWithPipeline2(currentDocument);
                    if (result != null && result.Succeeded) {
                        Process.Start(
                            Directory.Exists(converter.ConversionParameters.OutputPath)
                            ? converter.ConversionParameters.OutputPath
                            : Path.GetDirectoryName(converter.ConversionParameters.OutputPath)
                        );
                    } else {
                        MessageBox.Show(result.UnknownErrorMessage, "Conversion failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                } else {
                    eventsHandler.onConversionCanceled();
                }

                //applicationObject.ActiveDocument.Save();
            }
            catch (Exception e) {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }
        }

        public void SaveAsDTBookXML(IRibbonControl control, ConversionParameters conversionIntegrationTestSettings = null)
        {
            try
            {
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null 
                    ? (IConversionEventsHandler) new GraphicalEventsHandler() 
                    : new SilentEventsHandler();

                Script pipelineScript = Directory.Exists(ConverterHelper.Pipeline2Path)
                    ? new WordToCleanedDtbook(eventsHandler) :
                    null;
                if (pipelineScript != null)
                {
                    pipelineScript.EventsHandler = eventsHandler;
                }
                ApplyScript(pipelineScript, eventsHandler, conversionIntegrationTestSettings);
                NotifyDonationRequest();
            }
            catch (Exception e)
            {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }
            
        }

        /// <summary>
        /// UI Call : request conversion of the current active document to DAISY book
        /// </summary>
        /// <param name="control"></param>
        public void SaveAsDAISY3(IRibbonControl control, ConversionParameters conversionIntegrationTestSettings = null) {
            try {
               // IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new GraphicalEventsHandler() : new SilentEventsHandler();
                //Script pipelineScript = control != null ? this.PostprocessingPipeline?.getScript(control.Tag) : null;
                
                Script pipelineScript = Directory.Exists(ConverterHelper.Pipeline2Path)
                    ? new WordToDaisy3(eventsHandler) :
                    null;
               
                
                if (pipelineScript != null)
                {
                    pipelineScript.EventsHandler = eventsHandler;
                }
                ApplyScript(pipelineScript, eventsHandler, conversionIntegrationTestSettings);
                NotifyDonationRequest();
            } catch (Exception e) {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }
            
        }

        /// <summary>
        /// UI Call : request conversion of the current active document to DAISY book
        /// </summary>
        /// <param name="control"></param>
        public void SaveAsDAISY202(IRibbonControl control, ConversionParameters conversionIntegrationTestSettings = null)
        {
            try {
                // IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new GraphicalEventsHandler() : new SilentEventsHandler();
                //Script pipelineScript = control != null ? this.PostprocessingPipeline?.getScript(control.Tag) : null;

                Script pipelineScript = Directory.Exists(ConverterHelper.Pipeline2Path)
                    ? new WordToDaisy202(eventsHandler) :
                    null;

                if (pipelineScript != null) {
                    pipelineScript.EventsHandler = eventsHandler;
                }
                ApplyScript(pipelineScript, eventsHandler, conversionIntegrationTestSettings);
                NotifyDonationRequest();
            }
            catch (Exception e) {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }

        }

        /// <summary>
        /// Experimental conversion to epub3
        /// </summary>
        /// <param name="control"></param>
        /// <param name="conversionIntegrationTestSettings"></param>
        public void SaveAsEPUB3(IRibbonControl control, ConversionParameters conversionIntegrationTestSettings = null) {
            try {
                //IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new GraphicalEventsHandler() : new SilentEventsHandler();

                Script pipelineScript = new WordToEpub3(eventsHandler);

                if (pipelineScript != null)
                {
                    pipelineScript.EventsHandler = eventsHandler;
                }
                ApplyScript(pipelineScript, eventsHandler, conversionIntegrationTestSettings);
                NotifyDonationRequest();
            } catch (Exception e) {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }

        }

        /// <summary>
        /// Experimental conversion to epub3
        /// </summary>
        /// <param name="control"></param>
        /// <param name="conversionIntegrationTestSettings"></param>
        public void SaveAsMP3(IRibbonControl control, ConversionParameters conversionIntegrationTestSettings = null)
        {
            try
            {
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new GraphicalEventsHandler() : new SilentEventsHandler();

                Script pipelineScript = new WordToMp3(eventsHandler);

                if (pipelineScript != null)
                {
                    pipelineScript.EventsHandler = eventsHandler;
                }
                ApplyScript(pipelineScript, eventsHandler, conversionIntegrationTestSettings);
            }
            catch (Exception e)
            {
                AddinLogger.Error(e);
                if (conversionIntegrationTestSettings == null) {
                    ExceptionReport report = new ExceptionReport(e);
                    report.ShowDialog();
                }

            }

        }

        #endregion

        #region Multiple documents conversion

        /// <summary>
        /// UI Call : requesting the conversion of a list of documents into a single DTBook XML or DAISY book
        /// </summary>
        /// <param name="control"></param>
        public void MutipleWordToDTBookXML(IRibbonControl control) {
            try {
                GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                Script pipelineScript = new DtbookCleaner(eventsHandler);

                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if(documentsPathes != null && documentsPathes.Count > 0) {
                    List<DocumentParameters> documents = new List<DocumentParameters>();
                   
                    foreach (string inputPath in documentsPathes) {
                        DocumentParameters subDoc = null;
                        try {
                            subDoc = converter.PreprocessDocument(inputPath);
                        } catch (Exception e) {
                            string errors = "Convertion aborted due to the following errors found while preprocessing " + inputPath + ":\r\n" + e.Message;
                            eventsHandler.onPreprocessingError(inputPath, new Exception(errors, e));
                        }
                        if (subDoc != null) {
                            documents.Add(subDoc);
                        } else {
                            // abort documents conversion
                            documents.Clear();
                            break;
                        }
                    }
                    if(documents.Count > 0) converter.ConvertWithPipeline2AndMerge(documents);
                }

                applicationObject.ActiveDocument.Save();
            } catch (Exception e) {
                ExceptionReport report = new ExceptionReport(e);
                report.Show();
            }
        }

        public void MutipleWordToDAISY3(IRibbonControl control)
        {
            try
            {
                GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                Script pipelineScript = new WordToDaisy3(eventsHandler);

                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if (documentsPathes != null && documentsPathes.Count > 0)
                {
                    List<DocumentParameters> documents = new List<DocumentParameters>();

                    foreach (string inputPath in documentsPathes)
                    {
                        DocumentParameters subDoc = null;
                        try
                        {
                            subDoc = converter.PreprocessDocument(inputPath);
                        }
                        catch (Exception e)
                        {
                            string errors = "Convertion aborted due to errors found while preprocessing " + inputPath;
                            eventsHandler.onPreprocessingError(inputPath, new Exception(errors,e));
                        }
                        if (subDoc != null)
                        {
                            documents.Add(subDoc);
                        }
                        else
                        {
                            // abort documents conversion
                            documents.Clear();
                            break;
                        }
                    }
                    if (documents.Count > 0) converter.ConvertWithPipeline2AndMerge(documents);
                }

                applicationObject.ActiveDocument.Save();
            }
            catch (Exception e)
            {
                ExceptionReport report = new ExceptionReport(e);
                report.Show();
            }
        }

        public void MutipleWordToEPUB3(IRibbonControl control)
        {
            try
            {
                GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                Script pipelineScript = new WordToDaisy3(eventsHandler);

                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if (documentsPathes != null && documentsPathes.Count > 0)
                {
                    List<DocumentParameters> documents = new List<DocumentParameters>();

                    foreach (string inputPath in documentsPathes)
                    {
                        DocumentParameters subDoc = null;
                        try
                        {
                            subDoc = converter.PreprocessDocument(inputPath);
                        }
                        catch (Exception e)
                        {
                            string errors = "Convertion aborted due to errors found while preprocessing " + inputPath;
                            eventsHandler.onPreprocessingError(inputPath, new Exception(errors, e));
                        }
                        if (subDoc != null)
                        {
                            documents.Add(subDoc);
                        }
                        else
                        {
                            // abort documents conversion
                            documents.Clear();
                            break;
                        }
                    }
                    if (documents.Count > 0) converter.ConvertWithPipeline2AndMerge(documents);
                }

                applicationObject.ActiveDocument.Save();
            }
            catch (Exception e)
            {
                ExceptionReport report = new ExceptionReport(e);
                report.Show();
            }
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

        #endregion

        #region Abbreviation and Acronyms

        /// <summary>
        /// Core Function for Managing Abbreviation
        /// </summary>
        /// <param name="control"></param>
        public void ManageAbbreviation(IRibbonControl control)
        {
            try
            {

                MSword.Document doc = this.applicationObject.ActiveDocument;
                if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                {
                    AliasesManagement form = new AliasesManagement(doc, AliasesManagement.AliasType.Abbreviation);
                    form.ShowDialog();
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
        /// Core Function for Managing Abbreviations
        /// </summary>
        /// <param name="control"></param>
        public void ManageAcronym(IRibbonControl control)
        {
            try
            {
                MSword.Document doc = this.applicationObject.ActiveDocument;
                if (doc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                {
                    AliasesManagement form = new AliasesManagement(doc, AliasesManagement.AliasType.Acronym);
                    form.ShowDialog();
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
        /// Core function to Mark Abbreviations
        /// </summary>
        /// <param name="control"></param>
        public void MarkAsAbbreviationUI(IRibbonControl control)
        {
            MSword.Document doc = this.applicationObject.ActiveDocument;

            if ((this.applicationObject.Selection.Start == this.applicationObject.Selection.End) || this.applicationObject.Selection.Text.Equals("\r"))
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
                Mark mrkForm = new Mark(doc, AliasesManagement.AliasType.Abbreviation);
                mrkForm.ShowDialog();
            }
        }

        /// <summary>
        /// Function to check whether the selected Bookmark is Abbreviation\Acronym
        /// </summary>
        /// <param name="select">Range of text</param>
        /// <returns></returns>
        public String ValidateBookMark(MSword.Selection select)
        {
            MSword.Document currentDoc = this.applicationObject.ActiveDocument;
            String bkmrkValue = "Nobookmarks";

            if (this.applicationObject.Selection.Bookmarks.Count > 0)
            {

                foreach (object item in currentDoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Range.Text.Trim() == this.applicationObject.Selection.Text.Trim())
                    {
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
        /// Core Function for Managing Acronyms
        /// </summary>
        /// <param name="control"></param>
        public void MarkAsAcronymUI(IRibbonControl control)
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
            else if (ValidateBookMark(this.applicationObject.Selection) == "Acrtrue")
            {
                MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (ValidateBookMark(this.applicationObject.Selection) == "Abbrtrue")
            {
                MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Mark mrkForm = new Mark(doc, AliasesManagement.AliasType.Acronym);
                mrkForm.ShowDialog();
            }
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

        #endregion

        #region FootNote actions
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

        #region Language actions 

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
}