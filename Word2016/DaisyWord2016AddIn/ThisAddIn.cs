using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;


using System.IO;
using System.Collections;
using System.Xml;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using MSword = Microsoft.Office.Interop.Word;
using Sonata.DaisyConverter.DaisyConverterLib;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO.Packaging;


namespace DaisyWord2016AddIn
{


	using Microsoft.Win32;
	using System.Text;
	using Word = Microsoft.Office.Interop.Word.InlineShape;
	using IConnectDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
	using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
	using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
	using ConnectIEnumETC = System.Runtime.InteropServices.ComTypes.IEnumFORMATETC;
	using COMException = System.Runtime.InteropServices.COMException;
	using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;
    using DaisyWord2007AddIn;

    /// <summary>
    /// Word addin default class extension <br/>
    /// This class handles events and actions on the Application and document side. <br/>
    /// UI actions are defined in the DaisyRibbon class. <br/>
    /// Most of the codebase from here is taken from the office2007 addin, but adapted to launch using the ThisAddIn Class <br/>
    /// (which is automatically instantiate by office according to 
    /// https://docs.microsoft.com/en-us/visualstudio/vsto/improving-the-performance-of-a-vsto-add-in?view=vs-2019)    
    /// </summary>
    public partial class ThisAddIn {

        #region Relationships for file search in a word "package" (docx file)
        const string coreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        const string coreNamespace = "http://purl.org/dc/elements/1.1/";
        const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        const string subDocRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";
        const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        #endregion

        PackageRelationship packRelationship = null;

        PackageRelationship relationship = null;

        XmlDocument docTemp, mergeXmlDoc, currentDocXml, validation_xml;
        
        String docFile, path_For_Xml, templateName, path_For_Pipeline;
        
        public ArrayList docValidation = new ArrayList();
        
        
        
        private bool showValidateTabBool = false;
        private bool showViewTabBool = false;

        ArrayList mergeDoclanguage, openSubdocs;
        
        DaisyWord2007AddIn.frmValidate2007 validateInput;
        
        ArrayList listmathML;
        
        Pipeline pipe;
        
        ToolStripMenuItem PipelineMenuItem = null;
        
        private ArrayList footntRefrns;
        
        Hashtable multipleMathMl = new Hashtable();
        
        Hashtable multipleOwnMathMl = new Hashtable();
        
        ArrayList langMergeDoc, notTranslatedDoc;
        

        private MSword.Application applicationObject;
        
        /// <summary>
        /// Ribbon class that handles UI actions
        /// </summary>
        private DaisyRibbon daisyRibbon;

        #region Office addin entry point and "destructor"
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Adding event handlers
            this.applicationObject = Globals.ThisAddIn.Application;
            this.applicationObject.DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(applicationObject_DocumentBeforeSave);
            this.applicationObject.DocumentOpen += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentOpenEventHandler(applicationObject_DocumentOpen);
            this.applicationObject.DocumentChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentChangeEventHandler(applicationObject_DocumentChange);
            this.applicationObject.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(applicationObject_DocumentBeforeClose);
            this.applicationObject.WindowDeactivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowDeactivateEventHandler(applicationObject_WindowDeactivate);

            // Preparing resources for addin execution
            
            string assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            
            // - write the dotx template
            string templateFile = Path.Combine(assemblyDir, "template2007.dotx");
            if (!File.Exists(templateFile)) {
                File.WriteAllBytes(templateFile, Properties.Resources.template2007);
            }

            // - write the docx manual
            string docxManualFile = Path.Combine(assemblyDir, "Daisy_Translator_Instruction_Manual_01_March_2011_v2_5.docx");
            if (!File.Exists(docxManualFile)) {
                File.WriteAllBytes(docxManualFile, Properties.Resources.Daisy_Translator_Instruction_Manual_01_March_2011_v2_5);
            }

            // - write help file
            string helpFile = Path.Combine(assemblyDir, "Help.chm");
            if (!File.Exists(helpFile)) {
                File.WriteAllBytes(helpFile, Properties.Resources.Help);
            }

            // - write dtbook manual
            string dtbookManualDir = Path.Combine(assemblyDir, "Save-as-instruction-manual");
            if (!Directory.Exists(dtbookManualDir)) {
                File.WriteAllBytes(Path.Combine(assemblyDir, "Save-as-instruction-manual.zip"), Properties.Resources.Save_as_instruction_manual);
                Package zip = Package.Open(
                    Path.Combine(assemblyDir, "Save-as-instruction-manual.zip"),
                    FileMode.Open,
                    FileAccess.Read);
                var zipArchiveProperty = zip.GetType().GetField(
                        "_zipArchive",
                        BindingFlags.Static
                        | BindingFlags.NonPublic
                        | BindingFlags.Instance);
                var zipArchive = zipArchiveProperty.GetValue(zip);
                var zipFileInfoDictionaryProperty = zipArchive.GetType().GetProperty(
                        "ZipFileInfoDictionary",
                        BindingFlags.Static
                        | BindingFlags.NonPublic
                        | BindingFlags.Instance);
                var zipFileInfoDictionary =
                    zipFileInfoDictionaryProperty.GetValue(zipArchive, null) as Hashtable;
                foreach (System.Collections.DictionaryEntry zipFileEntry in zipFileInfoDictionary) {
                    string filePathInZip = zipFileEntry.Key.ToString();
                    if (!filePathInZip.EndsWith("/")) { // File is not a folder
                        // Create the directory tree if it does not exists
                        string targetFile = Path.Combine(assemblyDir, filePathInZip);
                        if (!Directory.Exists(Path.GetDirectoryName(targetFile))) {
                            Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                        }
                        var getStream = zipFileEntry.Value.GetType().GetMethod("GetStream", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        Stream fileStream = (Stream)getStream.Invoke(zipFileEntry.Value, new object[] { FileMode.Open, FileAccess.Read });
                        Stream outputFileStream = new FileStream(targetFile, FileMode.Create, FileAccess.Write);
                        fileStream.CopyTo(outputFileStream);
                    } else {
                        string targetDir = Path.Combine(assemblyDir, filePathInZip);
                        if (!Directory.Exists(targetDir)) {
                            Directory.CreateDirectory(targetDir);
                        }
                    }
                }
            }

            // - write and unzip Pipeline lite
            string pipelineDir = Path.Combine(assemblyDir, "pipeline-lite-ms");
            if (!Directory.Exists(pipelineDir)) {
                File.WriteAllBytes(Path.Combine(assemblyDir, "pipeline.zip"), Properties.Resources.pipeline_lite_ms);
                Package zip = Package.Open(
                    Path.Combine(assemblyDir, "pipeline.zip"),
                    FileMode.Open,
                    FileAccess.Read);
                var zipArchiveProperty = zip.GetType().GetField("_zipArchive",
                                        BindingFlags.Static
                                        | BindingFlags.NonPublic
                                        | BindingFlags.Instance);
                var zipArchive = zipArchiveProperty.GetValue(zip);
                var zipFileInfoDictionaryProperty =
                    zipArchive.GetType().GetProperty(
                        "ZipFileInfoDictionary",
                        BindingFlags.Static
                        | BindingFlags.NonPublic
                        | BindingFlags.Instance);
                var zipFileInfoDictionary =
                    zipFileInfoDictionaryProperty.GetValue(zipArchive, null) as Hashtable;
                foreach (System.Collections.DictionaryEntry zipFileEntry in zipFileInfoDictionary) {
                    string filePathInZip = zipFileEntry.Key.ToString();
                    if (!filePathInZip.EndsWith("/")) { // File is not a folder
                                                        // Create the directory tree
                        string targetFile = Path.Combine(assemblyDir, filePathInZip);
                        if (!Directory.Exists(Path.GetDirectoryName(targetFile))) {
                            Directory.CreateDirectory(Path.GetDirectoryName(targetFile));
                        }
                        var getStream = zipFileEntry.Value.GetType().GetMethod("GetStream", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        Stream fileStream = (Stream)getStream.Invoke(zipFileEntry.Value, new object[] { FileMode.Open, FileAccess.Read });
                        Stream outputFileStream = new FileStream(targetFile, FileMode.Create, FileAccess.Write);
                        fileStream.CopyTo(outputFileStream);
                    } else {
                        string targetDir = Path.Combine(assemblyDir, filePathInZip);
                        if (!Directory.Exists(targetDir)) {
                            Directory.CreateDirectory(targetDir);
                        }
                    }
                }
            }
            

        }
        
        /// <summary>
        /// Function to do changes in Look and Feel of UI
        /// </summary>
        void applicationObject_DocumentChange() {
            if (this.applicationObject.Documents.Count >= 1) {
                MSword.Document currentDoc = this.applicationObject.ActiveDocument;
                templateName = (currentDoc.get_AttachedTemplate() as MSword.Template).Name;
                CheckforAttchments();
                showValidateTabBool = false;
                daisyRibbon?.getLoadedUI().InvalidateControl("toggleValidate");
            }
        }
        //Envent handling deactivation of word document
        void applicationObject_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn) {

        }
        /// <summary>
        ///Core Function to check the Daisy Styles in the current Document
        /// </summary>
        /// <param name="Doc">Current Document</param>
        void applicationObject_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc) {
            MSword.Document doc = this.applicationObject.ActiveDocument;
            templateName = (doc.get_AttachedTemplate() as MSword.Template).Name;
            CheckforAttchments();
        }


        /// <summary>
        /// Function to check whether the Daisy Styles are new or not
        /// </summary>
        public void CheckforAttchments() {
            MSword.Document doc = this.applicationObject.ActiveDocument;
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
                daisyRibbon?.getLoadedUI().InvalidateControl("Button7");
            } else {
                showViewTabBool = false;
                daisyRibbon?.getLoadedUI().InvalidateControl("Button7");
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


        void applicationObject_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel) {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            if (this.daisyRibbon == null) this.daisyRibbon = new DaisyRibbon(this);
            return this.daisyRibbon;
        }


        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
