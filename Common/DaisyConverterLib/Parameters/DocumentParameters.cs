using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion
{

    /// <summary>
    /// Document specific parameters
    /// 
    /// </summary>
    public class DocumentParameters
    {
        /// <summary>
        /// Document-specific parameters
        /// </summary>
        /// <param name="inputPath">The original file used as input</param>
        public DocumentParameters(string inputPath)
        {
            InputPath = inputPath;
            CopyPath = ConverterHelper.GetTempPath(inputPath, ".docx");

            ObjectShapes = new List<string>();
            ImageIds = new List<string>();
            InlineShapes = new List<string>();
            InlineIds = new List<string>();
            SubDocumentsToConvert = new List<DocumentParameters>();
            Creator = "";
            Title = "";
            Publisher = "";
            HasRevisions = false;
        }
        
        /// <summary>
        /// Extract propertes from the copy. <br/>
        /// needs to be called by document preprocessor after it has created the copy but before
        /// the copy has been opened back for preprocessing or after preprocessing is finished and document is closed.
        /// (else the file could be not readable as being opened by another process)
        /// </summary>
        public void updatePropertiesFromCopy()
        {
            try {
                if (CopyPath != null && File.Exists(CopyPath)) using (
                        Package pack = Package.Open(CopyPath, FileMode.Open, FileAccess.Read)
                ) {
                    Creator = pack.PackageProperties.Creator;
                    Title = pack.PackageProperties.Title;
                    // Search Title in content xml if not declared in properties
                    if (Title == "") {
                        XmlDocument wordContentDocument = getFirstDocumentFromRelationship(
                            pack,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                        );
                        if (wordContentDocument != null) {
                            bool titleFlag = false;
                            string styleVal = "";
                            string msgConcat = "";
                            NameTable nt = new NameTable();
                            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                            nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                            try {
                                XmlNodeList getParagraph = wordContentDocument.SelectNodes("//w:body/w:p/w:pPr/w:pStyle", nsManager);
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
                                            titleFlag = true;
                                            break;
                                        }
                                        if (titleFlag) {
                                            break;
                                        }
                                    }
                                    if (titleFlag) {
                                        break;
                                    }
                                }
                                Title = msgConcat;
                            }
                            catch (Exception e) {
                                AddinLogger.Warning("An exception was raised while searching Title in content", e);
                            }
                            // while we are checking the content, also check for revisions
                            XmlNodeList listDel = wordContentDocument.SelectNodes("//w:del", nsManager);
                            XmlNodeList listIns = wordContentDocument.SelectNodes("//w:ins", nsManager);

                            HasRevisions = listDel.Count > 0 || listIns.Count > 0;
                        }
                    }
                    // Search publisher in extended properties xml
                    XmlDocument propertiesXml = getFirstDocumentFromRelationship(
                        pack,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                    );
                    if (propertiesXml != null) {
                        NameTable nt = new NameTable();
                        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                        nsManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                        XmlNodeList node = propertiesXml.GetElementsByTagName("Company");
                        if (node != null && node.Count > 0)
                            Publisher = node.Item(0).InnerText;
                    }


                };
            }
            catch (Exception e2) {
                AddinLogger.Warning("An exception was raised while parsing document properties", e2);
            }
        }


        public enum DocType
        {
            Simple,
            Master,
            Sub
        }

        public string Creator { get; private set; }

        public string Title { get; private set; }

        public string Publisher { get; private set; }


        /// <summary>
        /// Word document type between :<br/>
        /// - Simple : the document is self contained<br/>
        /// - Master : the document refers to subdocuments<br/>
        /// - Sub : the document is refered by another document <br/>
        /// Note : <br/>
        /// - Master Document will have subdocuments <br/>
        /// - SubDocument will have a resource ID <br/>
        /// - Simple will have neither
        /// </summary>
        public DocType Type
        {
            get
            {
                if (this.ResourceId != null)
                {
                    return DocType.Sub;
                }
                else if (this.SubDocumentsToConvert.Count > 0)
                {
                    return DocType.Master;
                }
                else return DocType.Simple;
            }
        }

        /// <summary>
        /// Original path/URL of the input
        /// </summary>
        public string InputPath { get; }

        /// <summary>
        /// Document copy to use for processing. <br/>
        /// Note : the copy is actually made by the DocumentPreprocessor class
        /// (that is specific to a word version as it uses word interop to save the copy)
        /// </summary>
        public string CopyPath { get; }

        /// <summary>
        /// Output Path of the document conversion
        /// Used as intermediate target for multiple conversion (batch or merged documents)
        /// where the conversion parameters output define the final target
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// If true, the input document is (re)opened in word during preprocessing, 
        /// when preprocessing starts and after the copy used for conversion is saved
        /// </summary>
        public bool ShowInputDocumentInWord { get; set; } = true;

        /// <summary>
        /// Pathes of the shapes images extracted during preprocessing
        /// </summary>
        public List<string> ObjectShapes { get; set; }

        /// <summary>
        /// Dictionnary of key => List of mathml equations (stored as string)
        /// </summary>
        public Dictionary<string, List<string>> MathMLMap { get; set; } = new Dictionary<string, List<string>>()
        {
            {"wdTextFrameStory", new List<string>() },
            {"wdFootnotesStory", new List<string>() },
            {"wdMainTextStory", new List<string>() },
        };

        /// <summary>
        /// Ids of the shapes extracted during preprocessing
        ///
        /// </summary>
        public List<string> ImageIds { get; set; }

        /// <summary>
        /// Pathes of the inline shapes images extracted during preprocessing
        /// </summary>
        public List<string> InlineShapes { get; set; }

        /// <summary>
        /// Ids of the inline shapes extracted during preprocessing
        /// </summary>
        public List<string> InlineIds { get; set; }


        /// <summary>
        /// Check if the current document as unaccepted revisions
        /// </summary>
        public bool HasRevisions { get; private set; }


        public bool TrackChanges = false;

        /// <summary>
        /// Sub documents referenced by the current document
        /// </summary>
        public bool HasSubDocuments = false;
        public List<DocumentParameters> SubDocumentsToConvert { get; set; }

        /// <summary>
        /// Resource ID of the document it it is a sub document contained in a Master document
        /// </summary>
        public string ResourceId { get; set; }

        public string GetInputFileNameWithoutExtension
        {
            get
            {
                int lastSeparatorIndex = InputPath.LastIndexOf('\\');
                // Special case : onedrive documents uses https based URL format with '/' as separator
                if (lastSeparatorIndex < 0)
                {
                    lastSeparatorIndex = InputPath.LastIndexOf('/');
                }
                if (lastSeparatorIndex < 0)
                { // no path separator found
                    return InputPath.Remove(InputPath.LastIndexOf('.'));
                }
                else
                {
                    string tempInput = InputPath.Substring(lastSeparatorIndex);
                    return tempInput.Remove(tempInput.LastIndexOf('.'));
                }
            }
        }

        /// <summary>
        /// Document parameters hash
        /// (To be used for the Daisy class used in xslt)
        /// </summary>
        public Hashtable ParametersHash
        {
            get
            {
                Hashtable parameters = new Hashtable
                {
                    { "TRACK", TrackChanges },
                    { "MasterSubFlag", (SubDocumentsToConvert != null && SubDocumentsToConvert.Count > 0) ? true : false }
                };


                return parameters;
            }
        }

        public string serialize()
        {
            return JsonConvert.SerializeObject(this);
        }

        /// <summary>
        /// Search and retrieve the first xml document matching a relationship url within an opened office package
        /// </summary>
        /// <param name="pack">the office package (.docx)</param>
        /// <param name="relationshipType">the url of the relationship associated with the searched document</param>
        /// <returns></returns>
        private static XmlDocument getFirstDocumentFromRelationship(Package pack, string relationshipType)
        {
            PackageRelationship packRelationship = null;
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(relationshipType)) {
                packRelationship = searchRelation;
                break;
            }
            if (packRelationship != null) {
                Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(mainPartxml.GetStream());
                return doc;
            }
            return null; // default to no document returned
        }



    }
}
