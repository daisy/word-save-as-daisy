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
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Reflection;
using System.Collections;
using System.IO.Packaging;
using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.CommandLineTool {
    /// <summary>Exception thrown when the file is not valid</summary>
    public class OoxValidatorException : Exception {
        public OoxValidatorException(String msg) : base(msg) { }
    }

    /// <summary>Check the validity of a docx file. Throw an OoxValidatorException if errors occurs</summary>
    public class OoxValidator {

        // namespaces 
        private static string OOX_DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // OOX special files
        private static string OOX_RELATIONSHIP_FILE = "_rels/.rels";

        // OOX relationship
        private static string OOX_DOCUMENT_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

        private XmlReaderSettings settings = null; // global settings to open the xml files
        private Report report;
        private Package reader;
        PackagePart relationshipPart = null;
        Stream str = null;
        String target = null;

        /// <summary>
        /// Initialize the validator
        /// </summary>
        public OoxValidator(Report report) {
            this.settings = new XmlReaderSettings();
            this.report = report;
        }


        /// <summary>
        /// Check the validity of an Office Open XML file.
        /// </summary>
        /// <param name="fileName">The path of the docx file.</param>
        public void validate(String fileName) {

            // 0. The file must exist and be a valid Package

            try {
                reader = Package.Open(fileName);
            } catch (Exception e) {
                throw new OoxValidatorException("Problem opening the docx file : " + e.Message);
            }

            // 1. _rels/.rels must be present and valid
            Stream relationShips = null;
            try {
                relationShips = GetEntry(reader, OOX_RELATIONSHIP_FILE);
            } catch (Exception) {
                throw new OoxValidatorException("The docx package must have a \"/_rels/.rels\" file");
            }
            this.validateXml(relationShips);

            // 2. _rel/.rels must contain a relationship of document.xml
            relationShips = GetEntry(reader, OOX_RELATIONSHIP_FILE);
            XmlReader r = XmlReader.Create(relationShips);
            String docTarget = null;
            while (r.Read() && docTarget == null) {
                if (r.NodeType == XmlNodeType.Element && r.GetAttribute("Type") == OOX_DOCUMENT_RELATIONSHIP_TYPE) {
                    docTarget = r.GetAttribute("Target");
                }
            }
            if (docTarget == null) {
                throw new OoxValidatorException(" document.xml relation not found in \"/_rels/.rels\"");
            }

            // 3. For each item in _rels/.rels
            relationShips = GetEntry(reader, OOX_RELATIONSHIP_FILE);
            r = XmlReader.Create(relationShips);
            while (r.Read()) {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship") {
                    String target = r.GetAttribute("Target");

                    // 3.1. The target item must exist in the package
                    Stream item = null;
                    try {
                        item = GetEntry(reader, target);
                    } catch (Exception) {
                        throw new OoxValidatorException("The file \"" + target + "\" is described in the \"/_rels/.rels\" file but does not exist in the package.");
                    }

                    // 3.2. A content type can be found in [Content_Types].xml file
                    String ct = this.FindContentType(reader, "/" + target);


                    //// 3.3. If it's an xml file, it has to be valid
                    if (ct.EndsWith("+xml")) {
                        this.validateXml(item);
                    }
                }
            }

            // Does a part relationship exist ?
            Stream partRel = null;
            String partDir = docTarget.Substring(0, docTarget.LastIndexOf("/"));
            String partRelPath = partDir + "/_rels/" + docTarget.Substring(docTarget.LastIndexOf("/") + 1) + ".rels";
            bool partRelExists = true;
            try {
                partRel = GetEntry(reader, partRelPath);
            } catch (Exception) {
                partRelExists = false;
            }

            if (partRelExists) {
                this.validateXml(partRel);

                // 4. For each item in /word/_rels/document.xml.rels
                partRel = GetEntry(reader, partRelPath);
                r = XmlReader.Create(partRel);
                target = null;

                while (r.Read()) {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship") {

                        if (r.GetAttribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml") {
                            //target = r.GetAttribute("Target").Replace("../", "");
                        } else if (r.GetAttribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink") {
                            //target = r.GetAttribute("Target").Replace("../", "");
                        } else if (r.GetAttribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" && r.GetAttribute("TargetMode") == "External") {
                            //target = r.GetAttribute("Target").Replace("../", "");
                        } else if (r.GetAttribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument" && r.GetAttribute("TargetMode") == "External") {
                            //target = r.GetAttribute("Target").Replace("../", "");
                        } else {
                            target = partDir + "/" + r.GetAttribute("Target");
                        }

                        // Is the target item exist in the package ?
                        Stream item = null;
                        bool fileExists = true;
                        try {
                            if (target != null) {
                                item = GetEntry(reader, target);
                            } else
                                fileExists = false;
                        } catch (Exception) {
                            fileExists = false;
                        }

                        if (fileExists) {
                            // 4.1. A content type can be found in [Content_Types].xml file
                            String ct = this.FindContentType(reader, "/" + target);

                            // 4.2. If it's an xml file, it has to be valid
                            if (ct.EndsWith("+xml")) {
                                this.validateXml(item);
                            }
                        }
                    }
                }
            }




            // 5. are all referenced relationships exist in the part relationship file
            // retrieve all ids referenced in the document

            Stream doc = GetEntry(reader, docTarget);
            r = XmlReader.Create(doc);

            ArrayList ids = new ArrayList();
            while (r.Read()) {
                if (r.NodeType == XmlNodeType.Element && r.GetAttribute("id", OOX_DOC_REL_NS) != null) {
                    if (!ids.Contains(r.GetAttribute("id", OOX_DOC_REL_NS)))
                        ids.Add(r.GetAttribute("id", OOX_DOC_REL_NS));
                }
            }

            // check if each id exists in the partRel file

            if (ids.Count != 0) {
                if (!partRelExists) {
                    throw new OoxValidatorException("Referenced id exist but no part relationship file found");
                }
                relationShips = GetEntry(reader, partRelPath);
                r = XmlReader.Create(relationShips);
                while (r.Read()) {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship") {
                        if (ids.Contains(r.GetAttribute("Id"))) {
                            ids.Remove(r.GetAttribute("Id"));
                        }
                    }
                }
                if (ids.Count != 0) {
                    throw new OoxValidatorException("One or more relationship id have not been found in the partRelationship file : " + ids[0]);
                }
            }
            reader.Close();
        }

        // validate xml stream
        private void validateXml(Stream xmlStream) {
            XmlReader r = XmlReader.Create(xmlStream, this.settings);
            while (r.Read()) ;
        }

        // find the content type of a part in the package
        private String FindContentType(Package reader, String target) {
            String extension = null;
            String contentType = null;
            if (target.IndexOf(".") != -1) {
                extension = target.Substring(target.IndexOf(".") + 1);
            }

            foreach (PackagePart searchRelation in reader.GetParts()) {
                relationshipPart = searchRelation;
                if (target == relationshipPart.Uri.ToString())
                    contentType = relationshipPart.ContentType.ToString();

            }

            if (contentType == null) {
                throw new OoxValidatorException("Content type not found for " + target);
            }
            return contentType;
        }

        public void ValidationHandler(object sender, ValidationEventArgs args) {
            throw new OoxValidatorException("XML Schema Validation error : " + args.Message);
        }

        public Stream GetEntry(Package pack, String name) {
            foreach (PackagePart searchRelation in pack.GetParts()) {
                String com = "/" + name;
                relationshipPart = searchRelation;
                if (com == relationshipPart.Uri.ToString())
                    str = relationshipPart.GetStream();
            }
            return str;

        }
    }
}
