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


using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using Daisy.SaveAsDAISY.Forms;
using Daisy.SaveAsDAISY.WPF;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;
using System.Xml;
using Converter = Daisy.SaveAsDAISY.Conversion.Converter;
using MSword = Microsoft.Office.Interop.Word;
using Script = Daisy.SaveAsDAISY.Conversion.Script;


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

               
                // https://stackoverflow.com/a/12768858
                //var formPreload = new Thread(() =>
                //{   //Force runtime to pre-load all resources for secondary windows so they will load as fast as possible
                //    new WPF.About();
                //    new WPF.ConversionParametersForm(null, null);
                //    new WPF.ConversionProgress();
                //    new WPF.SettingsForm();
                //    new WPF.Metadata();
                //    new WPF.ExceptionReport();
                //});
                //formPreload.SetApartmentState(ApartmentState.STA);
                //formPreload.Priority = ThreadPriority.Lowest;//We don't want prefetching to delay showing of primary window
                //formPreload.Start();


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
            // kill the embedded engine if running
            Engine.StopEmbeddedEngine();
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
            // Launch the pipeline in the background to start conversions asap
            try {
                if (ConverterSettings.Instance.UseDAISYPipelineApp) {
                    Engine.StartDAISYPipelineApp();
                } else {
                    Engine.StopEmbeddedEngine();
                    Engine.StartEmbeddedEngine();
                }
            } catch(Exception e) {
                AddinLogger.Error(e);
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
                if (Directory.Exists(ConverterHelper.EmbeddedEnginePath))
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
            { "DocumentationTabMenu", "speaker.jpg" },
            { "DocumentMetadataTabButton", "version.png" }
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
            Daisy.SaveAsDAISY.WPF.SettingsForm settings = new Daisy.SaveAsDAISY.WPF.SettingsForm();
            settings.ShowDialog();

            if (ConverterSettings.Instance.UseDAISYPipelineApp) {
                Engine.StopEmbeddedEngine();
                Engine.StartDAISYPipelineApp();
            } else {
                Engine.StartEmbeddedEngine();
            }

            
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
        }

        private Dictionary<string, Conversion.DocumentProperties> DocumentPropertiesCache = new Dictionary<string, Conversion.DocumentProperties>();

        /// <summary>
        /// Access document properties
        /// </summary>
        /// <param name="control"></param>
        public void DocumentMetadataUI(IRibbonControl control)
        {
            
            try {
                var dispatch = Dispatcher.CurrentDispatcher;
                var preprocessor = new DocumentPreprocessor(applicationObject);
                if (!DocumentPropertiesCache.ContainsKey(this.applicationObject.ActiveDocument.FullName)) {
                    var progress = new WPF.ConversionProgress();
                    dispatch.Invoke(() =>
                    {
                        progress.Show();
                        progress.InitializeProgress("Loading document metadata ...", 1, 1);
                    });
                    
                    object doc = this.applicationObject.ActiveDocument;
                    var propsTemp = preprocessor.loadDocumentParameters(ref doc);
                    progress.Close();
                    DocumentPropertiesCache.Add(this.applicationObject.ActiveDocument.FullName, propsTemp);
                }
                var props = DocumentPropertiesCache[this.applicationObject.ActiveDocument.FullName];
                dispatch.Invoke(() =>
                {
                    Metadata metadata = new Metadata(
                        props,
                        this.applicationObject.ActiveProtectedViewWindow == null
                    );

                    metadata.ShowDialog();
                    DocumentPropertiesCache[this.applicationObject.ActiveDocument.FullName] = metadata.UpdatedDocumentData;
                    if (metadata.MetadataUpdated) {
                        var progress = new WPF.ConversionProgress();
                        dispatch.Invoke(() =>
                        {
                            progress.Show();
                            progress.InitializeProgress("Updating document metadata ...", 1, 1);
                        });
                        object doc = this.applicationObject.ActiveDocument;
                        //var preprocessor = new DocumentPreprocessor(applicationObject);
                        preprocessor.updateDocumentMetadata(ref doc, metadata.UpdatedDocumentData);
                        progress.Close();
                    }
                });
                
            }
            catch (Exception e) {
                AddinLogger.Error(e);
                MessageBox.Show("The following error occured during metadata reading or saving, please contact the SaveAsDAISY team :\r\n" +e.Message, "Document Metadata error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Single document conversion



        /// <summary>
        /// New version of the export method
        /// </summary>
        /// <param name="pipelineScript"></param>
        /// <param name="eventsHandler"></param>
        public async void ApplyScript(Script pipelineScript, IConversionEventsHandler eventsHandler, ConversionParameters conversionIntegrationTestSettings = null)
        {
            Dispatcher uiThread = Dispatcher.CurrentDispatcher;
            ConversionResult result = ConversionResult.Success();
            try {
                eventsHandler.onFeedbackMessageReceived(this, new DaisyEventArgs("Starting conversion using " + pipelineScript.Name));
                object doc = this.applicationObject.ActiveDocument;
                DocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                Converter converter = new WPFConverter(preprocess, conversion, (WPFEventsHandler)eventsHandler);
                Exception catchedException = null;
                Conversion.DocumentProperties currentDocument = await System.Threading.Tasks.Task.Run(() =>
                {
                    // need to change preprocess to NOT do right now the shapes extraction
                    try {
                        return converter.AnalyzeDocument(this.applicationObject.ActiveDocument.FullName);
                    }
                    catch (Exception e) {
                        AddinLogger.Error(e);
                        catchedException = e;
                        return null;
                    }
                });
                if(catchedException != null) {
                    if (conversionIntegrationTestSettings == null) {
                        WPF.ExceptionReport report = new WPF.ExceptionReport(catchedException);
                        report.ShowDialog();
                    }
                }
                else if (currentDocument != null) {
                    
                    try {
                        WPF.ConversionParametersForm form = new WPF.ConversionParametersForm(
                        preprocess,
                        ref doc,
                        conversion,
                        currentDocument
                    );
                        if (form.ShowDialog()) {
                            currentDocument = form.DocumentProps;
                            converter.ConversionParameters = form.UpdatedConversionParameters;
                            
                            DirectoryInfo finalOutput = new DirectoryInfo(
                                Path.Combine(
                                    converter.ConversionParameters.OutputPath,
                                string.Format(
                                    "{0}_{2}_{1}",
                                    Path.GetFileNameWithoutExtension(currentDocument.InputPath),
                                    DateTime.Now.ToString("yyyyMMddHHmmssffff"),
                                    pipelineScript.Name
                                )
                            ));
                            // Remove and recreate result folder
                            // Since the DaisyToEpub3 requires output folder to be empty
                            if (finalOutput.Exists) {
                                finalOutput.Delete(true);
                                finalOutput.Create();
                            }
                            converter.ConversionParameters.OutputPath = finalOutput.FullName;
                            if (converter.ConversionParameters.PipelineScript.Parameters.ContainsKey("output")) {
                                converter.ConversionParameters.PipelineScript.Parameters["output"].Value = finalOutput.FullName;
                            }

                            // TODO add a chain of cancellation to stop processing if user cancel at any step

                            eventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Saving metadata in the document ..."));
                            preprocess.updateDocumentMetadata(ref doc, currentDocument);
                            var task = System.Threading.Tasks.Task.Run(() =>
                            {
                                try {
                                    var convresult = converter.ConvertWithPipeline2(currentDocument);
                                    System.Diagnostics.Process.Start(finalOutput.FullName);
                                    return convresult;
                                }
                                catch (JobException jex) {
                                    return ConversionResult.Fail(jex);
                                }
                                catch (Exception e) {
                                    AddinLogger.Error(e);
                                    catchedException = e;
                                    return ConversionResult.Fail(e);
                                }
                            });
                            
                            result = await task;

                        } else {
                            result = ConversionResult.Cancel();
                        }
                    }
                    catch (Exception e) {
                        AddinLogger.Error(e);
                        if (conversionIntegrationTestSettings == null) {
                            WPF.ExceptionReport report = new WPF.ExceptionReport(e);
                            report.ShowDialog();
                        }
                    }
                }
            } 
            catch (Exception e) {
                    AddinLogger.Error(e);
                    if (conversionIntegrationTestSettings == null) {
                        WPF.ExceptionReport report = new WPF.ExceptionReport(e);
                        report.ShowDialog();
                    }
            } finally {
                if (result.Canceled) {
                    eventsHandler.onConversionCanceled();
                } else if (result.Succeeded) {
                    eventsHandler.onConversionSuccess();
                } else if (result.Failed) {
                    MessageBox.Show(result.ErrorMessage, "Conversion failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

        public void ApplyScriptV1(Script pipelineScript, IConversionEventsHandler eventsHandler, ConversionParameters conversionIntegrationTestSettings = null)
        {
            try {
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(applicationObject);
                Converter converter = null;
                if (conversionIntegrationTestSettings != null) {
                    ConversionParameters conversion = conversionIntegrationTestSettings.withParameter("Version", this.applicationObject.Version);
                    converter = new Converter(preprocess, conversion, eventsHandler);
                } else {
                    ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                    converter = new GraphicalConverter(preprocess, conversion, (GraphicalEventsHandler)eventsHandler);
                }

                Conversion.DocumentProperties currentDocument = converter.AnalyzeDocument(this.applicationObject.ActiveDocument.FullName);
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
                        MessageBox.Show(result.ErrorMessage, "ConversionParametersForm failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    ? (IConversionEventsHandler) new WPFEventsHandler() 
                    : new SilentEventsHandler();
                Script pipelineScript = Directory.Exists(ConverterHelper.EmbeddedEnginePath)
                    //? new WordToCleanedDtbook(eventsHandler) :
                    ? new WordToDtbook(eventsHandler) :
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
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new WPFEventsHandler() : new SilentEventsHandler();
                //Script pipelineScript = control != null ? this.PostprocessingPipeline?.getScript(control.Tag) : null;
                
                Script pipelineScript = Directory.Exists(ConverterHelper.EmbeddedEnginePath)
                    ? new WordToDAISY3(eventsHandler) :
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
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new WPFEventsHandler() : new SilentEventsHandler();
                //Script pipelineScript = control != null ? this.PostprocessingPipeline?.getScript(control.Tag) : null;

                Script pipelineScript = Directory.Exists(ConverterHelper.EmbeddedEnginePath)
                    ? new WordToDAISY202(eventsHandler) :
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
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new WPFEventsHandler() : new SilentEventsHandler();

                Script pipelineScript = new WordToEPUB3(eventsHandler);

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
                IConversionEventsHandler eventsHandler = conversionIntegrationTestSettings == null ? (IConversionEventsHandler)new WPFEventsHandler() : new SilentEventsHandler();

                Script pipelineScript = new WordToMP3(eventsHandler);

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
                GraphicalConverter converter = new GraphicalConverter(preprocess, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if(documentsPathes != null && documentsPathes.Count > 0) {
                    List<Conversion.DocumentProperties> documents = new List<Conversion.DocumentProperties>();
                   
                    foreach (string inputPath in documentsPathes) {
                        Conversion.DocumentProperties subDoc = null;
                        try {
                            subDoc = converter.AnalyzeDocument(inputPath);
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
                Script pipelineScript = new WordToDAISY3(eventsHandler);

                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                GraphicalConverter converter = new GraphicalConverter(preprocess, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if (documentsPathes != null && documentsPathes.Count > 0)
                {
                    List<Conversion.DocumentProperties> documents = new List<Conversion.DocumentProperties>();

                    foreach (string inputPath in documentsPathes)
                    {
                        Conversion.DocumentProperties subDoc = null;
                        try
                        {
                            subDoc = converter.AnalyzeDocument(inputPath);
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
                Script pipelineScript = new WordToDAISY3(eventsHandler);

                ConversionParameters conversion = new ConversionParameters(this.applicationObject.Version, pipelineScript);
                GraphicalConverter converter = new GraphicalConverter(preprocess, conversion, eventsHandler);
                // Note : the current form for multiple also include conversion settings update
                List<string> documentsPathes = converter.requestUserDocumentsList();
                if (documentsPathes != null && documentsPathes.Count > 0)
                {
                    List<Conversion.DocumentProperties> documents = new List<Conversion.DocumentProperties>();

                    foreach (string inputPath in documentsPathes)
                    {
                        Conversion.DocumentProperties subDoc = null;
                        try
                        {
                            subDoc = converter.AnalyzeDocument(inputPath);
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
                    ManageAbbreviationsForm form = new ManageAbbreviationsForm(doc);
                    form.ShowDialog();
                    //AliasesManagement form = new AliasesManagement(doc, AliasesManagement.AliasType.Abbreviation);
                    //form.ShowDialog();
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
                    ManageAcronymsForm form = new ManageAcronymsForm(doc);
                    form.ShowDialog();
                    //AliasesManagement form = new AliasesManagement(doc, AliasesManagement.AliasType.Acronym);
                    //form.ShowDialog();
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
            else if (HasAbreviationOrAcronymBookmark(this.applicationObject.Selection) == BookmarkType.Abbreviation)
            {
                MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (HasAbreviationOrAcronymBookmark(this.applicationObject.Selection) == BookmarkType.Acronym)
            {
                MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                NewAbreviationForm newAbreviationForm = new NewAbreviationForm(doc);
                newAbreviationForm.ShowDialog();
                //Mark mrkForm = new Mark(doc, AliasesManagement.AliasType.Abbreviation);
                //mrkForm.ShowDialog();
            }
        }

        public enum BookmarkType
        {
            None, Abbreviation, Acronym
        }

        /// <summary>
        /// Function to check whether the selected Bookmark is Abbreviation\Acronym
        /// </summary>
        /// <param name="select">Range of text</param>
        /// <returns></returns>
        public BookmarkType HasAbreviationOrAcronymBookmark(MSword.Selection select)
        {
            MSword.Document currentDoc = this.applicationObject.ActiveDocument;

            if (this.applicationObject.Selection.Bookmarks.Count > 0)
            {
                foreach (object item in currentDoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Range.Text.Trim() == this.applicationObject.Selection.Text.Trim())
                    {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
                            return BookmarkType.Abbreviation;
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                           return BookmarkType.Acronym;
                    }
                }
            }

            return BookmarkType.None;
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
            else if (HasAbreviationOrAcronymBookmark(this.applicationObject.Selection) == BookmarkType.Acronym)
            {
                MessageBox.Show(addinLib.GetString("AcronymAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (HasAbreviationOrAcronymBookmark(this.applicationObject.Selection) == BookmarkType.Abbreviation)
            {
                MessageBox.Show(addinLib.GetString("AbbreviationAlready"), addinLib.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                NewAcronymForm newAcronymForm = new NewAcronymForm(doc);
                newAcronymForm.ShowDialog();
                //Mark mrkForm = new Mark(doc, AliasesManagement.AliasType.Acronym);
                //mrkForm.ShowDialog();
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