using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion
{

    /// <summary>
    /// Document specific parameters
    /// 
    /// </summary>
    public class DocumentProperties : ICloneable
    {
        /// <summary>
        /// Document-specific parameters
        /// </summary>
        /// <param name="inputPath">The original file used as input</param>
        public DocumentProperties(string inputPath)
        {
            InputPath = inputPath;
            CopyPath = ConverterHelper.GetTempPath(inputPath, ".docx");

            ObjectShapes = new List<string>();
            ImageIds = new List<string>();
            InlineShapes = new List<string>();
            InlineIds = new List<string>();
            SubDocumentsToConvert = new List<DocumentProperties>();
            Languages = new List<string>();
            HasRevisions = false;
        }
        
        /// <summary>
        /// Copy constructor
        /// </summary>
        /// <param name="toClone"></param>
        private DocumentProperties(DocumentProperties toClone)
        {
            InputPath = toClone.InputPath;
            CopyPath = toClone.CopyPath;

            ObjectShapes = toClone.ObjectShapes.ToList();
            ImageIds = toClone.ImageIds.ToList();
            InlineShapes = toClone.InlineShapes.ToList();
            InlineIds = toClone.InlineIds.ToList();
            SubDocumentsToConvert = toClone.SubDocumentsToConvert.ToList();
            Languages = toClone.Languages.ToList();
            HasRevisions = toClone.HasRevisions;

            Title = toClone.Title;
            Subtitle = toClone.Subtitle;
            Author = toClone.Author;
            Date = toClone.Date;
            Contributor = toClone.Contributor;
            Publisher = toClone.Publisher;
            Rights = toClone.Rights;
            Identifier = toClone.Identifier;
            IdentifierScheme = toClone.IdentifierScheme;
            SourceOfPagination = toClone.SourceOfPagination;
            SourceDate = toClone.SourceDate;
            Summary = toClone.Summary;
            Subject = toClone.Subject;
            AccessibilitySummary = toClone.AccessibilitySummary;

        }

        public List<string> Languages { get; set; }


        // Richard use caracter check to determine the script
        // taking from https://learn.microsoft.com/en-us/dotnet/standard/base-types/character-classes-in-regular-expressions)
        // instead of the shorthand declared in .net
        // (goal is to adapt that in xslts after for conversion language evaluation)
        // EastAsia codepoints 
        /*
         * 1100 - 11FF IsHangulJamo 
         * 1720 - 173F IsHanunoo 
         * 3040 - 309F IsHiragana  
         * 30A0 - 30FF IsKatakana 
         * 3100 - 312F IsBopomofo 
         * 3130 - 318F IsHangulCompatibilityJamo 
         * 31A0 - 31BF IsBopomofoExtended 
         * 31F0 - 31FF IsKatakanaPhoneticExtensions 
         * 4DC0 - 4DFF IsYijingHexagramSymbols 
         * A000 - A48F IsYiSyllables 
         * A490 - A4CF IsYiRadicals 
         * AC00 - D7AF IsHangulSyllables 
         */
        private static readonly string EastAsiaPattern = "[\\[\\]\\(\\)\\{\\}=\\+\\-_;\\.~\\$\\*\\%\\&\"',0123456789!\\?@#:<>\\/\\|]*[" +
            "\u1100-\u11FF" +
            "\u1720-\u173F" +
            "\u3040-\u309F" +
            "\u30A0-\u30FF" +
            "\u3100-\u312F" +
            "\u3130-\u318F" +
            "\u31A0-\u31BF" +
            "\u31F0-\u31FF" +
            "\u4DC0-\u4DFF" +
            "\u4E00-\u9FFF" +
            "\uA000-\uA48F" +
            "\uA490-\uA4CF" +
            "\uAC00-\uD7AF" +
        "]+.*";
        private static readonly Regex IsEastAsia = new Regex(EastAsiaPattern, RegexOptions.Compiled);
        // Bidi codepoints
        /*
         * 0590 - 05FF IsHebrew
         * 0600 - 06FF IsArabic
         * 0700 - 074F IsSyriac
         * 0780 - 07BF IsThaana
         * 0980 - 09FF IsBengali
         * 0900 - 097F IsDevanagari
         * 0A00 - 0A7F IsGurmukhi
         * 0A80 - 0AFF IsGujarati
         * 0B00 - 0B7F IsOriya
         * 0B80 - 0BFF IsTamil
         * 0C00 - 0C7F IsTelugu
         * 0C80 - 0CFF IsKannada
         * 0D00 - 0D7F IsMalayalam
         * 0D80 - 0DFF IsSinhala
         * 0E00 - 0E7F IsThai
         * 0E80 - 0EFF IsLao
         * 0F00 - 0FFF IsTibetan
         * 1000 - 109F IsMyanmar
         * 10A0 - 10FF IsGeorgian
         * FB50 - FDFF IsArabicPresentationForms-A
         * FE70 - FEFF IsArabicPresentationForms-B
         */
        private static readonly string BidiPattern = "[\\[\\]\\(\\)\\{\\}=\\+\\-_;\\.~\\$\\*\\%\\&\"',0123456789!\\?@#:<>\\/\\|]*[" +
            "\u0590-\u05FF" +
            "\u0600-\u06FF" +
            "\u0700-\u074F" +
            "\u0780-\u07BF" +
            "\u0980-\u09FF" +
            "\u0900-\u097F" +
            "\u0A00-\u0A7F" +
            "\u0A80-\u0AFF" +
            "\u0B00-\u0B7F" +
            "\u0B80-\u0BFF" +
            "\u0C00-\u0C7F" +
            "\u0C80-\u0CFF" +
            "\u0D00-\u0D7F" +
            "\u0D80-\u0DFF" +
            "\u0E00-\u0E7F" +
            "\u0E80-\u0EFF" +
            "\u0F00-\u0FFF" +
            "\u1000-\u109F" +
            "\u10A0-\u10FF" +
            "\uFB50-\uFDFF" +
            "\uFE70-\uFEFF" +
        "]+.*";
        private static readonly Regex IsBidi = new Regex(BidiPattern, RegexOptions.Compiled);

        //public static List<string> GetLanguagesFromDocx(Package pack)
        //{
        //    List<string> languages = new List<string>();
        //    // Word xml  
        //    XmlDocument wordContentDocument = getFirstDocumentFromRelationshipOrUri(
        //            pack,
        //            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        //        );
        //        XmlDocument stylesDocument = getFirstDocumentFromRelationshipOrUri(
        //            pack,
        //            "word/styles.xml"
        //        );
        //        NameTable nt = new NameTable();
        //        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
        //        nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        //        // Search publisher in extended properties xml
        //        XmlDocument propertiesXml = getFirstDocumentFromRelationshipOrUri(
        //            pack,
        //            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
        //        );
        //        // TODO extract list of languages and sort them by their count of presence in document
        //        // the count is obtained by checking each runner in word and counting each language
        //        // in it
        //        // NP 2024/06/12 : Richard had a VB code used in WordToEpub to evaluate more accurately the language of a run
        //        // I try to adapt the logic here
                
        //        List<int> languagesCount = new List<int>();
        //        if (stylesDocument != null && wordContentDocument != null) {
        //            string defaultLatin = "";
        //            string defaultEastAsia = "";
        //            string defaultComplex = "";
        //            //get default languages for the document
        //            XmlNodeList defaultRunnerlanguages = stylesDocument.SelectNodes(
        //                "//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang",
        //                nsManager
        //            );
        //            if (defaultRunnerlanguages.Count > 0) {
        //                defaultLatin = defaultRunnerlanguages[0].Attributes["w:val"].Value;
        //                if (defaultRunnerlanguages[0].Attributes["w:eastAsia"] != null) {
        //                    defaultEastAsia = defaultRunnerlanguages[0].Attributes["w:eastAsia"].Value;
        //                } else {
        //                    defaultEastAsia = defaultLatin;
        //                }
        //                if (defaultRunnerlanguages[0].Attributes["w:bidi"] != null) {
        //                    defaultComplex = defaultRunnerlanguages[0].Attributes["w:bidi"].Value;
        //                } else {
        //                    defaultComplex = defaultLatin;
        //                }
        //            }


        //            // get languages for style "Normal" style
        //            // According to Richard's code :
        //            // "If Normal has a language setting then this should be used as the default rather than the default para style" 
        //            XmlNodeList NormalStylelanguages = stylesDocument.SelectNodes(
        //                "//w:styles/w:style[@w:type='paragraph' and @w:styleId='Normal']/w:rPr/w:lang",
        //                nsManager
        //            );
        //            if (NormalStylelanguages.Count > 0) {
        //                if (NormalStylelanguages[0].Attributes["w:val"] != null) {
        //                    defaultLatin = NormalStylelanguages[0].Attributes["w:val"].Value;
        //                }
        //                if (NormalStylelanguages[0].Attributes["w:eastAsia"] != null) {
        //                    defaultEastAsia = NormalStylelanguages[0].Attributes["w:eastAsia"].Value;
        //                }
        //                if (NormalStylelanguages[0].Attributes["w:bidi"] != null) {
        //                    defaultComplex = NormalStylelanguages[0].Attributes["w:bidi"].Value;
        //                }
        //            }

        //            XmlNodeList defaultCharacterStyleLang = stylesDocument.SelectNodes(
        //                "//w:styles/w:style[w:type='character' and w:default='1']/w:rPr/w:lang",
        //                nsManager
        //            );

        //            // for each paragraph in document
        //            XmlNodeList paragraph = wordContentDocument.SelectNodes(
        //                "//w:p",
        //                nsManager
        //            );
        //            foreach (XmlNode paragraphNode in paragraph) {
        //                // default paragraph langs from default
        //                string paragraphLatin = defaultLatin;
        //                string paragraphEastAsia = defaultEastAsia;
        //                string paragraphComplex = defaultComplex;
        //                // Check style-specific languages

        //                XmlNodeList styleId = paragraphNode.SelectNodes("./w:pPr/w:pStyle/@w:val", nsManager);
        //                if (styleId.Count > 0) {
        //                    string styleVal = styleId[0].Value;
        //                    //character
        //                    // 
        //                    XmlNodeList styleLanguages = stylesDocument.SelectNodes(
        //                        $"//w:styles/w:style[@w:type='paragraph' and @w:styleId='{styleVal}']/w:rPr/w:lang",
        //                        nsManager
        //                    );
        //                    // Paragraph style has a language setting
        //                    if (styleLanguages.Count > 0) {
        //                        if (styleLanguages[0].Attributes["w:val"] != null) {
        //                            paragraphLatin = styleLanguages[0].Attributes["w:val"].Value;
        //                        }
        //                        if (styleLanguages[0].Attributes["w:eastAsia"] != null) {
        //                            paragraphEastAsia = styleLanguages[0].Attributes["w:eastAsia"].Value;
        //                        }
        //                        if (styleLanguages[0].Attributes["w:bidi"] != null) {
        //                            paragraphComplex = styleLanguages[0].Attributes["w:bidi"].Value;
        //                        }
        //                    }
        //                }
        //                XmlNodeList paragraphLanguages = paragraphNode.SelectNodes("./w:pPr/w:rPr/w:lang", nsManager);
        //                if (paragraphLanguages.Count > 0) {
        //                    // Parapgraph has a language setting for its runners
        //                    if (paragraphLanguages[0].Attributes["w:val"] != null) {
        //                        paragraphLatin = paragraphLanguages[0].Attributes["w:val"].Value;
        //                    }
        //                    if (paragraphLanguages[0].Attributes["w:eastAsia"] != null) {
        //                        paragraphEastAsia = paragraphLanguages[0].Attributes["w:eastAsia"].Value;
        //                    }
        //                    if (paragraphLanguages[0].Attributes["w:bidi"] != null) {
        //                        paragraphComplex = paragraphLanguages[0].Attributes["w:bidi"].Value;
        //                    }
        //                }

        //                XmlNodeList runners = paragraphNode.SelectNodes("./w:r", nsManager);
        //                foreach (XmlNode run in runners) {
        //                    // Default to the paragraph languages
        //                    string runLatin = paragraphLatin;
        //                    string runEastAsia = paragraphEastAsia;
        //                    string runComplex = paragraphComplex;
        //                    string runScript = "Latin";
        //                    string runText = run.InnerText.Trim();

        //                    if (IsEastAsia.IsMatch(runText)) {
        //                        runScript = "EastAsia";
        //                    } else if (IsBidi.IsMatch(runText)) {
        //                        runScript = "Complex";
        //                    }

        //                    XmlNode langNode = null;
        //                    XmlNodeList runLanguages = run.SelectNodes("./w:rPr/w:lang", nsManager);
        //                    if (runLanguages.Count > 0) {
        //                        // Check languages defined in runner properties
        //                        if (runLanguages[0].Attributes["w:val"] != null) {
        //                            runLatin = runLanguages[0].Attributes["w:val"].Value;
        //                        }
        //                        if (runLanguages[0].Attributes["w:eastAsia"] != null) {
        //                            runEastAsia = runLanguages[0].Attributes["w:eastAsia"].Value;
        //                        }
        //                        if (runLanguages[0].Attributes["w:bidi"] != null) {
        //                            runComplex = runLanguages[0].Attributes["w:bidi"].Value;
        //                        }
        //                    } else {
        //                        // check for runner styles
        //                        // <w:rPr> <w:rStyle w:val="TestCharacterStyle" />
        //                        XmlNodeList runStyleId = run.SelectNodes("./w:rPr/w:rStyle/@w:val", nsManager);
        //                        if (runStyleId.Count > 0) {
        //                            XmlAttribute xmlNode = (XmlAttribute)runStyleId[0];
        //                            string styleVal = xmlNode.Value;
        //                            XmlNodeList styleLanguages = stylesDocument.SelectNodes(
        //                                $"//w:styles/w:style[@w:type='character' and @w:styleId='{styleVal}']/w:rPr/w:lang",
        //                                nsManager
        //                            );
        //                            if (styleLanguages.Count > 0) {
        //                                if (styleLanguages[0].Attributes["w:val"] != null) {
        //                                    runLatin = styleLanguages[0].Attributes["w:val"].Value;
        //                                }
        //                                if (styleLanguages[0].Attributes["w:eastAsia"] != null) {
        //                                    runEastAsia = styleLanguages[0].Attributes["w:eastAsia"].Value;
        //                                }
        //                                if (styleLanguages[0].Attributes["w:bidi"] != null) {
        //                                    runComplex = styleLanguages[0].Attributes["w:bidi"].Value;
        //                                }
        //                            }
        //                        } else {
        //                            // Check if default character style has a language associated
        //                            if (defaultCharacterStyleLang.Count > 0) {
        //                                if (defaultCharacterStyleLang[0].Attributes["w:val"] != null) {
        //                                    runLatin = defaultCharacterStyleLang[0].Attributes["w:val"].Value;
        //                                }
        //                                if (defaultCharacterStyleLang[0].Attributes["w:eastAsia"] != null) {
        //                                    runEastAsia = defaultCharacterStyleLang[0].Attributes["w:eastAsia"].Value;
        //                                }
        //                                if (defaultCharacterStyleLang[0].Attributes["w:bidi"] != null) {
        //                                    runComplex = defaultCharacterStyleLang[0].Attributes["w:bidi"].Value;
        //                                }
        //                            }
        //                        } // Else the paragraph styling is used
        //                    }
        //                    // Default to latin
        //                    string langToAdd = runLatin;
        //                    if (runScript == "EastAsia") {
        //                        langToAdd = runEastAsia;
        //                    } else if (runScript == "Complex") {
        //                        langToAdd = runComplex;
        //                    }

        //                    int langId = languages.IndexOf(langToAdd);
        //                    if (langId == -1) {
        //                        languages.Add(langToAdd);
        //                        languagesCount.Add(runText.Length);
        //                    } else {
        //                        languagesCount[langId] += runText.Length;
        //                    }

        //                }
        //            }
        //        }
        //        if (languages.Count > 0) {
        //            // sort languages based on their associated counter
        //            // ( descending order )
        //            languages = languages
        //                .Select((x, i) => (x, i))
        //                .OrderByDescending(t => languagesCount[t.i])
        //                .Select(t => t.x)
        //                .ToList();
        //        }
        //}

        /// <summary>
        /// Extract propertes from the copy. <br/>
        /// needs to be called by document preprocessor after it has created the copy but before
        /// the copy has been opened back for preprocessing or after preprocessing is finished and document is closed.
        /// (else the file could be not readable as being opened by another process)
        /// </summary>
        public void updatePropertiesFromCopy()
        {
            try {
                if (CopyPath != null && File.Exists(CopyPath)) {
                    using (
                        Package pack = Package.Open(CopyPath, FileMode.Open, FileAccess.Read)
                    ) {
                        Author = pack.PackageProperties.Creator;
                        Title = pack.PackageProperties.Title;
                        // Word xml  
                        XmlDocument wordContentDocument = getFirstDocumentFromRelationshipOrUri(
                            pack,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                        );
                        XmlDocument stylesDocument = getFirstDocumentFromRelationshipOrUri(
                            pack,
                            "word/styles.xml"
                        );
                        NameTable nt = new NameTable();
                        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                        nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                        // Search Title in content xml if not declared in properties
                        if (Title == "" && wordContentDocument != null) {
                            bool titleFlag = false;
                            string styleVal = "";
                            string msgConcat = "";
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
                        // Search publisher in extended properties xml
                        XmlDocument propertiesXml = getFirstDocumentFromRelationshipOrUri(
                            pack,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
                        );
                        if (propertiesXml != null) {
                            nsManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                            XmlNodeList node = propertiesXml.GetElementsByTagName("Company");
                            if (node != null && node.Count > 0)
                                Publisher = node.Item(0).InnerText;
                        }
                        // TODO extract list of languages and sort them by their count of presence in document
                        // the count is obtained by checking each runner in word and counting each language
                        // in it
                        // NP 2024/06/12 : Richard had a VB code used in WordToEpub to evaluate more accurately the language of a run
                        // I try to adapt the logic here
                        Languages = new List<string>();
                        List<int> languagesCount = new List<int>();
                        if (stylesDocument != null && wordContentDocument != null) {
                            string defaultLatin = "";
                            string defaultEastAsia = "";
                            string defaultComplex = "";
                            //get default languages for the document
                            XmlNodeList defaultRunnerlanguages = stylesDocument.SelectNodes(
                                "//w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:lang",
                                nsManager
                            );
                            if (defaultRunnerlanguages.Count > 0) {
                                defaultLatin = defaultRunnerlanguages[0].Attributes["w:val"].Value;
                                if(defaultRunnerlanguages[0].Attributes["w:eastAsia"] != null) {
                                    defaultEastAsia = defaultRunnerlanguages[0].Attributes["w:eastAsia"].Value;
                                } else {
                                    defaultEastAsia = defaultLatin;
                                }
                                if (defaultRunnerlanguages[0].Attributes["w:bidi"] != null) {
                                    defaultComplex = defaultRunnerlanguages[0].Attributes["w:bidi"].Value;
                                } else {
                                    defaultComplex = defaultLatin;
                                }
                            }


                            // get languages for style "Normal" style
                            // According to Richard's code :
                            // "If Normal has a language setting then this should be used as the default rather than the default para style" 
                            XmlNodeList NormalStylelanguages = stylesDocument.SelectNodes(
                                "//w:styles/w:style[@w:type='paragraph' and @w:styleId='Normal']/w:rPr/w:lang",
                                nsManager
                            );
                            if (NormalStylelanguages.Count > 0) {
                                if(NormalStylelanguages[0].Attributes["w:val"] != null) {
                                    defaultLatin = NormalStylelanguages[0].Attributes["w:val"].Value;
                                }
                                if (NormalStylelanguages[0].Attributes["w:eastAsia"] != null) {
                                    defaultEastAsia = NormalStylelanguages[0].Attributes["w:eastAsia"].Value;
                                }
                                if (NormalStylelanguages[0].Attributes["w:bidi"] != null) {
                                    defaultComplex = NormalStylelanguages[0].Attributes["w:bidi"].Value;
                                }
                            }

                            XmlNodeList defaultCharacterStyleLang = stylesDocument.SelectNodes(
                                "//w:styles/w:style[w:type='character' and w:default='1']/w:rPr/w:lang",
                                nsManager
                            );

                            // for each paragraph in document
                            XmlNodeList paragraph = wordContentDocument.SelectNodes(
                                "//w:p",
                                nsManager
                            );
                            foreach ( XmlNode paragraphNode in paragraph ) {
                                // default paragraph langs from default
                                string paragraphLatin = defaultLatin;
                                string paragraphEastAsia = defaultEastAsia;
                                string paragraphComplex = defaultComplex;
                                // Check style-specific languages
                                
                                XmlNodeList styleId = paragraphNode.SelectNodes("./w:pPr/w:pStyle/@w:val", nsManager);
                                if (styleId.Count > 0) {
                                    string styleVal = styleId[0].Value;
                                    //character
                                    // 
                                    XmlNodeList styleLanguages = stylesDocument.SelectNodes(
                                        $"//w:styles/w:style[@w:type='paragraph' and @w:styleId='{styleVal}']/w:rPr/w:lang",
                                        nsManager
                                    );
                                    // Paragraph style has a language setting
                                    if (styleLanguages.Count > 0) {
                                        if (styleLanguages[0].Attributes["w:val"] != null) {
                                            paragraphLatin = styleLanguages[0].Attributes["w:val"].Value;
                                        }
                                        if (styleLanguages[0].Attributes["w:eastAsia"] != null) {
                                            paragraphEastAsia = styleLanguages[0].Attributes["w:eastAsia"].Value;
                                        }
                                        if (styleLanguages[0].Attributes["w:bidi"] != null) {
                                            paragraphComplex = styleLanguages[0].Attributes["w:bidi"].Value;
                                        }
                                    }
                                }
                                XmlNodeList paragraphLanguages = paragraphNode.SelectNodes("./w:pPr/w:rPr/w:lang", nsManager);
                                if(paragraphLanguages.Count > 0) {
                                    // Parapgraph has a language setting for its runners
                                    if (paragraphLanguages[0].Attributes["w:val"] != null) {
                                        paragraphLatin = paragraphLanguages[0].Attributes["w:val"].Value;
                                    }
                                    if (paragraphLanguages[0].Attributes["w:eastAsia"] != null) {
                                        paragraphEastAsia = paragraphLanguages[0].Attributes["w:eastAsia"].Value;
                                    }
                                    if (paragraphLanguages[0].Attributes["w:bidi"] != null) {
                                        paragraphComplex = paragraphLanguages[0].Attributes["w:bidi"].Value;
                                    }
                                }

                                XmlNodeList runners = paragraphNode.SelectNodes("./w:r", nsManager);
                                foreach (XmlNode run in runners) {
                                    // Default to the paragraph languages
                                    string runLatin = paragraphLatin;
                                    string runEastAsia = paragraphEastAsia;
                                    string runComplex = paragraphComplex;
                                    string runScript = "Latin";
                                    string runText = run.InnerText.Trim();
                                    
                                    if (IsEastAsia.IsMatch(runText)) {
                                        runScript = "EastAsia";
                                    } else if (IsBidi.IsMatch(runText)) {
                                        runScript = "Complex";
                                    }

                                    XmlNode langNode = null;
                                    XmlNodeList runLanguages = run.SelectNodes("./w:rPr/w:lang", nsManager);
                                    if (runLanguages.Count > 0) {
                                        // Check languages defined in runner properties
                                        if (runLanguages[0].Attributes["w:val"] != null) {
                                            runLatin = runLanguages[0].Attributes["w:val"].Value;
                                        }
                                        if (runLanguages[0].Attributes["w:eastAsia"] != null) {
                                            runEastAsia = runLanguages[0].Attributes["w:eastAsia"].Value;
                                        }
                                        if (runLanguages[0].Attributes["w:bidi"] != null) {
                                            runComplex = runLanguages[0].Attributes["w:bidi"].Value;
                                        }
                                    } else {
                                        // check for runner styles
                                        // <w:rPr> <w:rStyle w:val="TestCharacterStyle" />
                                        XmlNodeList runStyleId = run.SelectNodes("./w:rPr/w:rStyle/@w:val", nsManager);
                                        if (runStyleId.Count > 0) {
                                            XmlAttribute xmlNode = (XmlAttribute)runStyleId[0];
                                            string styleVal = xmlNode.Value;
                                            XmlNodeList styleLanguages = stylesDocument.SelectNodes(
                                                $"//w:styles/w:style[@w:type='character' and @w:styleId='{styleVal}']/w:rPr/w:lang",
                                                nsManager
                                            );
                                            if(styleLanguages.Count > 0) {
                                                if (styleLanguages[0].Attributes["w:val"] != null) {
                                                    runLatin = styleLanguages[0].Attributes["w:val"].Value;
                                                }
                                                if (styleLanguages[0].Attributes["w:eastAsia"] != null) {
                                                    runEastAsia = styleLanguages[0].Attributes["w:eastAsia"].Value;
                                                }
                                                if (styleLanguages[0].Attributes["w:bidi"] != null) {
                                                    runComplex = styleLanguages[0].Attributes["w:bidi"].Value;
                                                }
                                            }
                                        } else {
                                            // Check if default character style has a language associated
                                            if (defaultCharacterStyleLang.Count > 0) {
                                                if (defaultCharacterStyleLang[0].Attributes["w:val"] != null) {
                                                    runLatin = defaultCharacterStyleLang[0].Attributes["w:val"].Value;
                                                }
                                                if (defaultCharacterStyleLang[0].Attributes["w:eastAsia"] != null) {
                                                    runEastAsia = defaultCharacterStyleLang[0].Attributes["w:eastAsia"].Value;
                                                }
                                                if (defaultCharacterStyleLang[0].Attributes["w:bidi"] != null) {
                                                    runComplex = defaultCharacterStyleLang[0].Attributes["w:bidi"].Value;
                                                }
                                            }
                                        } // Else the paragraph styling is used
                                    }
                                    // Default to latin
                                    string langToAdd = runLatin;
                                    if(runScript == "EastAsia") {
                                        langToAdd = runEastAsia;
                                    } else if (runScript == "Complex") {
                                        langToAdd = runComplex;
                                    }

                                    int langId = Languages.IndexOf(langToAdd);
                                    if (langId == -1) {
                                        Languages.Add(langToAdd);
                                        languagesCount.Add(runText.Length);
                                    } else {
                                        languagesCount[langId] += runText.Length;
                                    }

                                }
                            }
                        }
                        if(Languages.Count > 0) {
                            // sort languages based on their associated counter
                            // ( descending order )
                            Languages = Languages
                                .Select((x, i) => (x, i))
                                .OrderByDescending(t => languagesCount[t.i])
                                .Select(t => t.x)
                                .ToList();
                        }

                        // retrieve custom properties
                        XmlDocument customProperties = getFirstDocumentFromRelationshipOrUri(
                            pack,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
                        );
                        if (customProperties != null) {
                            nsManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
                            var nodes1 = customProperties.GetElementsByTagName("property");
                            foreach (XmlNode property in customProperties.GetElementsByTagName("property")) {
                                if(property.Attributes["name"] == null) {
                                    continue;
                                }
                                switch (property.Attributes["name"].Value) {
                                    case "Contributor":
                                        this.Contributor = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Subtitle":
                                        this.Subtitle = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Publisher":
                                        this.Publisher = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? this.Publisher;
                                        break;
                                    case "Rights":
                                        this.Rights = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "ISBN":
                                        this.IdentifierScheme = "ISBN";
                                        this.Identifier = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Identifier":
                                        this.Identifier = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "IdentifierScheme":
                                        this.IdentifierScheme = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Date":
                                        this.Date = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Description":
                                        this.Summary = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "Source":
                                        this.SourceOfPagination = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "AccessibilitySummary":
                                        this.AccessibilitySummary = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                    case "SourceDate":
                                        this.SourceDate = property.SelectNodes("./vt:lpwstr", nsManager)[0]?.InnerText ?? "";
                                        break;
                                }
                            }
                        }
                    }
                }
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

        #region Document metadata

        public string Title { get; set; } = "";
        public string Subtitle { get; set; } = "";
        public string Author { get; set; } = "";
        public string Date { get; set; } = "";
        public string Contributor { get; set; } = "";
        public string Publisher { get; set; } = "";
        public string Rights { get; set; } = "";
        public string Identifier { get; set; } = "";
        public string IdentifierScheme { get; set; } = "ISBN";
        public string SourceOfPagination { get; set; } = "";
        public string SourceDate { get; set; } = "";
        public string Summary { get; set; } = "";
        public string Subject { get; set; } = "";
        public string AccessibilitySummary { get; set; } = "";


        /* 
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
   xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ContentTypeId">
       <vt:lpwstr>0x0101008B12E3ADF2E9EC46AE1D62AE160ABD91</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MTWinEqns">
       <vt:bool>true</vt:bool>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="_AdHocReviewCycleID">
       <vt:i4>-162943600</vt:i4>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="_NewReviewCycle">
       <vt:lpwstr></vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="_EmailSubject">
       <vt:lpwstr>Save As DAISY</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="7" name="_AuthorEmail">
       <vt:lpwstr>prashant.rv@gmail.com</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="8" name="_AuthorEmailDisplayName">
       <vt:lpwstr>prashant.rv@gmail.com</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="9" name="_ReviewingToolsShownOnce">
       <vt:lpwstr></vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="10" name="Contributor">
       <vt:lpwstr>Nicolas Pavie</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="11" name="Subtitle">
       <vt:lpwstr>Subtitle test</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="12" name="Publisher">
       <vt:lpwstr>Association Valentin Hauy</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="13" name="Rights">
       <vt:lpwstr>Rights field test</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="14" name="ISBN">
       <vt:lpwstr>9781234567890</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="15" name="Date">
       <vt:lpwstr>2025-06-11</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="16" name="Description">
       <vt:lpwstr>I don't know</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="17" name="Source">
       <vt:lpwstr>The Word document was used as the source for page numbers.  (NP: content here was autogenerated by word to epub)</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="18" name="AccessibilitySummary">
       <vt:lpwstr>The publication contains structural and page navigation. Alt text is included for
           images if this was present in source Word document. (NP: content here was autogenerated by word to epub)</vt:lpwstr>
   </property>
   <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="19" name="SourceDate">
       <vt:lpwstr></vt:lpwstr>
   </property>
</Properties>

    */

        #endregion

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
        public bool HasRevisions { get; set; }


        public bool AcceptRevisions = false;

        /// <summary>
        /// Sub documents referenced by the current document
        /// </summary>
        public bool HasSubDocuments = false;
        public List<DocumentProperties> SubDocumentsToConvert { get; set; }

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
        public Dictionary<string,object> ParametersValues
        {
            get
            {
                Dictionary<string, object> parameters = new Dictionary<string, object>
                {
                    {"title", Title},
                    {"subtitle", Subtitle},
                    {"author", Author},
                    {"date", Date},
                    {"contributor", Contributor},
                    {"publisher", Publisher},
                    {"rights", Rights},
                    {"uid", Identifier},
                    {"identifierscheme", IdentifierScheme},
                    {"source", SourceOfPagination},
                    {"sourcedate", SourceDate},
                    {"summary", Summary},
                    {"subject", Subject},
                    { "accessibility-summary", AccessibilitySummary},
                    { "accept-revisions", AcceptRevisions },
                    { "MasterSub", (SubDocumentsToConvert != null && SubDocumentsToConvert.Count > 0) ? true : false },
                    { "language", Languages != null && Languages.Count > 0 ? Languages[0] : "" }
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
        private static XmlDocument getFirstDocumentFromRelationshipOrUri(Package pack, string relationshipTypeOrUri)
        {
            PackageRelationship packRelationship = null;
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(relationshipTypeOrUri)) {
                packRelationship = searchRelation;
                break;
            }
            if (packRelationship != null) {
                Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(mainPartxml.GetStream());
                return doc;
            } else {
                try {
                    Uri testSource = new Uri("/", UriKind.Relative);
                    Uri testTarget = new Uri(relationshipTypeOrUri, UriKind.Relative);
                    Uri partUri = PackUriHelper.ResolvePartUri(
                        testSource,
                        testTarget
                    );
                    PackagePart mainPartxml = pack.GetPart(partUri);
                    XmlDocument doc = new XmlDocument();
                    doc.Load(mainPartxml.GetStream());
                    return doc;
                } catch {
                    return null;
                }
            }
        }

        public object Clone()
        {
            return new DocumentProperties(this);
        }
    }
}
