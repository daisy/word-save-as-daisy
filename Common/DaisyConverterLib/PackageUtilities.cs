using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {
    public class PackageUtilities {
        #region Conversion preprocessing utilities
        /// <summary>
        /// Function to get the Title of the Current Document
        /// </summary>
        /// <returns>Title</returns>
        public static string DocPropTitle(string docFile) {
            const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            PackageRelationship packRelationship = null;

            int titleFlag = 0;
            string styleVal = "";
            string msgConcat = "";
            Package pack;
            pack = Package.Open(docFile, FileMode.Open, FileAccess.ReadWrite);
            string strTitle = pack.PackageProperties.Title;
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
        public static String DocPropCreator(string docFile) {

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
        public static String DocPropPublish(string docFile) {

            const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
            const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
            PackageRelationship packRelationship = null;

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

        /// <summary>
        /// Function to test if a word Document is open
        /// </summary>
        /// <returns>message</returns>
        public static bool documentIsOpen(string documentPath) {
            try {
                Package pack;
                pack = Package.Open(documentPath, FileMode.Open, FileAccess.ReadWrite);
                pack.Close();
            } catch {
                return true;
            }
            return false;
        }
    }
}
