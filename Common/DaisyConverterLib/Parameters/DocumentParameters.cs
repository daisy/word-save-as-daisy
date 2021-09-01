using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Document specific parameters
    /// 
    /// </summary>
    public class DocumentParameters {

        public enum DocType {
            Simple,
            Master,
            Sub
        }


        public DocumentParameters(string inputPath) {
            this.InputPath = inputPath;
            
            ListMathMl = new Hashtable();
            ObjectShapes = new List<string>();
            ImageIds = new List<string>();
            InlineShapes = new List<string>();
            InlineIds = new List<string>();
            SubDocumentsToConvert = new List<DocumentParameters>();
        }


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
        public DocType Type {
            get {
                if(this.ResourceId != null) {
                    return DocType.Sub;
                } else if (this.SubDocumentsToConvert.Count > 0) {
                    return DocType.Master;
                } else return  DocType.Simple;
            }
        }

        /// <summary>
        /// Original path/URL of the input
        /// </summary>
        public string InputPath { get; set; }

        /// <summary>
        /// Document copy to use for processing
        /// </summary>
        public string CopyPath { get; set; }

        /// <summary>
        /// Output Path of the document conversion
        /// Used as intermediate target for multiple conversion (batch or merged documents)
        /// where the conversion parameters output define the final target
        /// </summary>
        public string OutputPath { get; set; }


        public List<string> ObjectShapes { get; set; }

        /// <summary>
        /// To be replaced by a list of string later on
        /// (Needs to update the DaisyClass object for that)
        /// </summary>
        public Hashtable ListMathMl { get; set; }
        public List<string> ImageIds { get; set; }
        public List<string> InlineShapes { get; set; }
        public List<string> InlineIds { get; set; }


        

        private bool? _hasRevisions = null;

        /// <summary>
        /// Check if the current document as unaccepted revisions
        /// </summary>
        public bool HasRevisions { get {
                if(!_hasRevisions.HasValue) {
                    const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
                    Package pack = Package.Open(this.CopyPath ?? this.InputPath, FileMode.Open, FileAccess.ReadWrite);
                    PackageRelationship packRelationship = null;
                    foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                        packRelationship = searchRelation;
                        break;
                    }

                    Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                    PackagePart mainPartxml = pack.GetPart(partUri);

                    XmlDocument XmlPackage = new XmlDocument();
                    XmlPackage.Load(mainPartxml.GetStream());
                    pack.Close();

                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("w", docNamespace);

                    XmlNodeList listDel = XmlPackage.SelectNodes("//w:del", nsManager);
                    XmlNodeList listIns = XmlPackage.SelectNodes("//w:ins", nsManager);

                    _hasRevisions = listDel.Count > 0 || listIns.Count > 0;
                }
                return _hasRevisions.Value;
            } }


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

        public string GetInputFileNameWithoutExtension {
            get {
                int lastSeparatorIndex = InputPath.LastIndexOf('\\');
                // Special case : onedrive documents uses https based URL format with '/' as separator
                if (lastSeparatorIndex < 0) {
                    lastSeparatorIndex = InputPath.LastIndexOf('/');
                }
                if (lastSeparatorIndex < 0) { // no path separator found
                    return InputPath.Remove(InputPath.LastIndexOf('.'));
                } else {
                    string tempInput = InputPath.Substring(lastSeparatorIndex);
                    return tempInput.Remove(tempInput.LastIndexOf('.'));
                }
            }
        }

        /// <summary>
        /// Parameters 
        /// </summary>
        public Hashtable ParametersHash { 
            get {
                Hashtable parameters = new Hashtable();

                parameters.Add("TRACK", !HasRevisions ? "NoTrack" : (TrackChanges ? "Yes" : "No"));
                parameters.Add("MasterSubFlag", !HasSubDocuments ? "NoMasterSub" : (SubDocumentsToConvert.Count > 0 ? "Yes" : "No"));
                
                return parameters;
            }
        }
    }
}
