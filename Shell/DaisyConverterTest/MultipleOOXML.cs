using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Collections;
using System.IO.Packaging;
using System.Reflection;
using Sonata.DaisyConverter.DaisyConverterLib;

namespace Sonata.DaisyConverter.CommandLineTool
{
    class MultipleOOXML
    {
        private Report report;
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        PackageRelationship relationship = null;
        ArrayList langMergeDoc, notTranslatedDoc;
        private OoxValidator ooxValidator = null;
        private static string error_MasterSub = "";
        XmlDocument mergeXmlDoc;
        private static string error;
        ArrayList mergeDocLanguage;
        ArrayList lostElements = new ArrayList();
        ArrayList fidilityLossMsg = new ArrayList();
        int subCount = 1;
        private string xslPath = null;
        ArrayList subList;
        String docName = "";
        int subDocFootNum;
        String cutData, error_Exception = "";
        String tempData = "";
        private static bool isValid;
        private ArrayList skipedPostProcessors = null;   // Post processors to skip (identified by their names)
        private Direction transformDirection = Direction.DocxToXml; // direction of conversion
        ArrayList MathList8879, MathList9573, MathListmathml;
        String errorText = "";
        Hashtable myLabel = new Hashtable();
        Hashtable listMathMl;

        /*Property which returns Fidility loss*/
        public ArrayList FidilityLoss
        {
            get
            {
                return fidilityLossMsg;
            }
        }

        /*Constructor*/
        public MultipleOOXML(Report report)
        {
            this.report = report;
            AddMathmlDtds();
            myLabel.Add("translation.oox2Daisy.commentReference", " Comment Reference is not translated");
            myLabel.Add("translation.oox2Daisy.ExternalImage", " External Images are not translated");
            myLabel.Add("translation.oox2Daisy.GeometricShapeElement", " Geometric shapes is is not translated");
            myLabel.Add("translation.oox2Daisy.ImageContent", " Image is Corrupted");
            myLabel.Add("translation.oox2Daisy.object", " Object is not translated");
            myLabel.Add("translation.oox2Daisy.sdtElement", " is not translated");
            myLabel.Add("translation.oox2Daisy.shapeElement", " Shapes are not translated");
            myLabel.Add("translation.oox2Daisy.UncoveredElement", " is not translated");
            myLabel.Add("translation.oox2Daisy.ImageCaption", "Caption above the Image is not translated");
            myLabel.Add("translation.oox2Daisy.TableCaption", "Caption below the Table is not translated");

        }

        /*Core Function to Copy Math Ml DTDS to destination folder*/
        public void AddMathmlDtds()
        {
            MathList8879 = new ArrayList();
            MathList9573 = new ArrayList();
            MathListmathml = new ArrayList();

            MathList8879.Add("isobox.ent");
            MathList8879.Add("isocyr1.ent");
            MathList8879.Add("isocyr2.ent");
            MathList8879.Add("isodia.ent");
            MathList8879.Add("isolat1.ent");
            MathList8879.Add("isolat2.ent");
            MathList8879.Add("isonum.ent");

            MathList8879.Add("isopub.ent");

            MathListmathml.Add("mmlalias.ent");
            MathListmathml.Add("mmlextra.ent");

            MathList9573.Add("isoamsa.ent");
            MathList9573.Add("isoamsb.ent");
            MathList9573.Add("isoamsc.ent");
            MathList9573.Add("isoamsn.ent");
            MathList9573.Add("isoamso.ent");
            MathList9573.Add("isoamsr.ent");
            MathList9573.Add("isogrk3.ent");
            MathList9573.Add("isomfrk.ent");
            MathList9573.Add("isomopf.ent");
            MathList9573.Add("isomscr.ent");
            MathList9573.Add("isotech.ent");
        }

        /*Property which return Count for sub documents*/
        public int FileCount
        {
            get
            {
                return subCount;
            }
        }

        /*Property which returns List of Sub documents*/
        public ArrayList DocListOwn
        {
            get
            {
                return subList;
            }
        }

        /*Fucntion returns Validation errors*/
        public string ValidationError
        {
            get
            {
                return error_Exception;
            }
        }

        /*Core Function to validate the input document*/
        private bool ValidateFile(string input)
        {
            try
            {
                if (this.ooxValidator == null)
                {
                    this.report.AddComment("Instanciating validator...");
                    this.ooxValidator = new OoxValidator(this.report);
                }
                this.ooxValidator.validate(input);
                this.report.AddComment("**********************");
                this.report.AddLog("", "Input file (" + input + ") is valid", Report.INFO_LEVEL);
                return true;
            }
            catch (OoxValidatorException e)
            {
                this.report.AddComment("**********************");
                this.report.AddLog("", "Input file (" + input + ") is invalid", Report.WARNING_LEVEL);
                this.report.AddLog("", e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }
            catch (Exception e)
            {
                this.report.AddLog("", "An unexpected exception occured when trying to validate " + input, Report.ERROR_LEVEL);
                this.report.AddLog("input", e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }

        }

        /* Function to Add  all the Messages to an Array */
        private void FeedbackMessageInterceptor(object sender, EventArgs e)
        {
            string message = ((DaisyEventArgs)e).Message;
            string messageValue = null;
            if (message.Contains("Cover Pages"))
                message = message.Replace("Cover Pages", "Cover Page");

            if (message.Contains("|"))
            {
                string[] messageKey = message.Split('|');
                int index = messageKey[0].IndexOf('%');
                // parameters substitution
                if (index > 0)
                {
                    string[] param = messageKey[0].Substring(index + 1).Split(new char[] { '%' });
                    foreach (DictionaryEntry myEntry in myLabel)
                    {
                        if (myEntry.Key.ToString().Equals(messageKey[0].Substring(0, index)))
                            messageValue = myEntry.Value.ToString();
                    }

                    if (messageValue != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            messageValue = messageValue.Replace("%" + (i + 1), param[i]);
                        }
                    }
                }
                else
                {
                    foreach (DictionaryEntry myEntry in myLabel)
                    {
                        if (myEntry.Key.ToString().Equals(messageKey[0]))
                            messageValue = myEntry.Value.ToString();
                    }

                }

                if (messageValue != null && !fidilityLossMsg.Contains(messageKey[1] + messageValue + " for " + Path.GetFileName(docName)))
                {
                    fidilityLossMsg.Add(messageKey[1] + messageValue + " for " + Path.GetFileName(docName));
                }
            }
            else
            {
                int index = message.IndexOf('%');
                // parameters substitution
                if (index > 0)
                {
                    string[] param = message.Substring(index + 1).Split(new char[] { '%' });
                    foreach (DictionaryEntry myEntry in myLabel)
                    {
                        if (myEntry.Key.ToString().Equals(message.Substring(0, index)))
                            messageValue = myEntry.Value.ToString();
                    }

                    if (messageValue != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            messageValue = messageValue.Replace("%" + (i + 1), param[i]);
                        }
                    }
                }
                else
                {
                    foreach (DictionaryEntry myEntry in myLabel)
                    {
                        if (myEntry.Key.ToString().Equals(message))
                            messageValue = myEntry.Value.ToString();
                    }
                }

                if (messageValue != null && !fidilityLossMsg.Contains(messageValue + " for " + Path.GetFileName(docName)))
                {
                    fidilityLossMsg.Add(messageValue + " for " + Path.GetFileName(docName));
                }
            }
        }


        #region Own Multiple OOXML

        /* Core Function to translate all Sub documents */
        public void MultipleOwnDocCore(String inputFile, String outputfilepath)
        {
            try
            {
                subList = new ArrayList();
                langMergeDoc = new ArrayList();
                notTranslatedDoc = new ArrayList();
                subList.Add(inputFile + "|Master");

                //OPening the Package of the Current Document
                Package packDoc;
                packDoc = Package.Open(inputFile, FileMode.Open, FileAccess.ReadWrite);

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
                                    fileName = Path.GetDirectoryName(inputFile) + "//" + fileName;
                                    if (File.Exists(fileName))
                                    {
                                        subList.Add(fileName + "|" + relationship.Id.ToString());
                                    }
                                }
                                else if (fileName.Contains("file"))
                                {
                                    fileName = fileName.Replace("file:///", "");
                                    if (File.Exists(fileName))
                                    {
                                        subList.Add(fileName + "|" + relationship.Id.ToString());
                                    }
                                }
                                else
                                {
                                    fileName = Path.GetDirectoryName(inputFile) + "\\" + Path.GetFileName(fileName);
                                    if (fileName.Contains("%20"))
                                        fileName = fileName.Replace("%20", " ");
                                    if (File.Exists(fileName))
                                    {
                                        subList.Add(fileName + "|" + relationship.Id.ToString());
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
                                    fileName = Path.GetDirectoryName(inputFile) + "//" + fileName;
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
                                    fileName = Path.GetDirectoryName(inputFile) + "\\" + Path.GetFileName(fileName);
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
            }
            catch (Exception e)
            {
                //MessageBox.Show(this.resourceManager.GetString("TranslationFailed") + "\n" + this.resourceManager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + validationErrorMsg + "\n" + "Problem is:" + "\n" + e.Message + "\n", "DAISY Translator", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /* Function to translate the current document along with sub documents */
        public void MultipleOwnDoc(String outputfilepath, ArrayList subList, Hashtable table)
        {
            try
            {
                error_MasterSub = "";
                mergeXmlDoc = new XmlDocument();
                mergeDocLanguage = new ArrayList();

                AbstractConverter converter = ConverterFactory.Instance(transformDirection);
                converter.ExternalResources = this.xslPath;
                converter.SkipedPostProcessors = this.skipedPostProcessors;
                converter.DirectTransform = transformDirection == Direction.DocxToXml;

                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    docName = splt[0];
                    converter.RemoveMessageListeners();
                    converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
                    converter.Transform(splt[0], null, null, null, true,"");
                }


                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    String outputFile = Path.GetDirectoryName(outputfilepath) + "\\" + Path.GetFileNameWithoutExtension(splt[0]) + ".xml";
                    String ridOutputFile = splt[1];
                    converter.Transform(splt[0], outputFile, table, listMathMl, true,"");
                    if (i == 0)
                    {
                        ReplaceData(outputFile);
                        mergeXmlDoc.Load(outputFile);

                        if (File.Exists(outputFile))
                        {
                            File.Delete(outputFile);
                        }
                    }
                    else
                    {
                        ReplaceData(outputFile);
                        MergeXml(outputFile, mergeXmlDoc, ridOutputFile, splt[0]);

                        if (File.Exists(outputFile))
                        {
                            File.Delete(outputFile);
                        }
                    }
                }
                SetPageNum(mergeXmlDoc);
                SetImage(mergeXmlDoc);
                SetLanguage(mergeXmlDoc);
                RemoveSubDoc(mergeXmlDoc);
                mergeXmlDoc.Save(outputfilepath);
                ReplaceData(outputfilepath, true);
                CopyDTDToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                CopyMATHToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                ValidateOutputFile(outputfilepath);
                ReplaceData(outputfilepath, false);
                if (File.Exists(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd"))
                {
                    File.Delete(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd");
                }
                DeleteMath(Path.GetDirectoryName(outputfilepath));
            }
            catch (Exception e)
            {
                //  error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + error_MasterSub + "\n" + "Problem is:" + "\n" + e.Message + "\n";

            }
        }

        /*Function to translate all Bunch of documents  selected by the user*/
        public void MultipleBatchDoc(String outputfilepath, ArrayList subList, Hashtable table)
        {
            try
            {
                error_MasterSub = "";
                mergeXmlDoc = new XmlDocument();
                mergeDocLanguage = new ArrayList();
                AbstractConverter converter = ConverterFactory.Instance(transformDirection);
                converter.ExternalResources = this.xslPath;
                converter.SkipedPostProcessors = this.skipedPostProcessors;
                converter.DirectTransform = transformDirection == Direction.DocxToXml;

                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    docName = splt[0];
                    converter.RemoveMessageListeners();
                    converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
                    converter.Transform(splt[0], null, null, null, true,"");
                }


                for (int i = 0; i < subList.Count; i++)
                {
                    bool validated = false;
                    validated = ValidateFile(subList[i].ToString());
                    if (validated)
                    {
                        String outputFile = Path.GetDirectoryName(outputfilepath) + "\\" + Path.GetFileNameWithoutExtension(subList[i].ToString()) + ".xml";
                        converter.Transform(subList[i].ToString(), outputFile, table, null, true,"");
                        if (i == 0)
                        {
                            ReplaceData(outputFile);
                            mergeXmlDoc.Load(outputFile);

                            if (File.Exists(outputFile))
                            {
                                File.Delete(outputFile);
                            }

                        }
                        else
                        {
                            ReplaceData(outputFile);
                            MergeXml(outputFile, mergeXmlDoc, subList[i].ToString());

                            if (File.Exists(outputFile))
                            {
                                File.Delete(outputFile);
                            }
                        }
                    }
                }

                SetPageNum(mergeXmlDoc);
                SetImage(mergeXmlDoc);
                SetLanguage(mergeXmlDoc);
                RemoveSubDoc(mergeXmlDoc);
                mergeXmlDoc.Save(outputfilepath);
                ReplaceData(outputfilepath, true);
                CopyDTDToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                CopyMATHToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                ValidateOutputFile(outputfilepath);
                ReplaceData(outputfilepath, false);
                if (File.Exists(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd"))
                {
                    File.Delete(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd");
                }
                DeleteMath(Path.GetDirectoryName(outputfilepath));

            }
            catch (Exception e)
            {
                error_Exception = "Translation failed." + "\n" + "Validation error found while translating the following documents" + "\n" + " \"" + Path.GetFileName(subList[0].ToString()) + "\"" + "\n" + "Validation error:" + "\n" + e.Message + "\n";
            }

        }



        #endregion

        #region MultipleOOXML Supproting functions

        /* Function checking for all the documents skipped during the Translation*/
        public String DocumentSkipped()
        {
            String message = "";
            if (notTranslatedDoc.Count != 0)
            {
                for (int i = 0; i < notTranslatedDoc.Count; i++)
                    message = message + Convert.ToString(i + 1) + ". " + notTranslatedDoc[i].ToString() + "\n";

                message = "\n\n" + "Files which are not in Word 2007 format are skipped during Translation:" + "\n" + message;
            }
            return message;
        }

        /*Function to validate the input document*/
        public void ValidateOutputFile(String outFile)
        {
            isValid = true;
            XmlTextReader xml = new XmlTextReader(outFile);
            XmlValidatingReader xsd = new XmlValidatingReader(xml);

            try
            {
                xsd.ValidationType = ValidationType.DTD;
                xsd.ValidationEventHandler += new ValidationEventHandler(MyValidationEventHandler);

                while (xsd.Read())
                {
                    errorText = xsd.ReadString();
                    if (errorText.Length > 100)
                        errorText = errorText.Substring(0, 100);
                }
                xsd.Close();

                Stream stream = null;
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (name.EndsWith("Shematron.xsl"))
                    {
                        stream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }

                XmlReader rdr = XmlReader.Create(stream);
                XPathDocument doc = new XPathDocument(outFile);

                XslCompiledTransform trans = new XslCompiledTransform(false);
                trans.Load(rdr);

                XmlTextWriter myWriter = new XmlTextWriter(Path.GetDirectoryName(outFile) + "\\report.txt", null);
                trans.Transform(doc, null, myWriter);

                myWriter.Close();
                rdr.Close();

                StreamReader reader = new StreamReader(Path.GetDirectoryName(outFile) + "\\report.txt");
                if (!reader.EndOfStream)
                {
                    error += reader.ReadToEnd();
                    isValid = false;
                }
                reader.Close();

                if (File.Exists(Path.GetDirectoryName(outFile) + "\\report.txt"))
                {
                    File.Delete(Path.GetDirectoryName(outFile) + "\\report.txt");
                }

                // Check whether the document is valid or invalid.
                if (isValid == false)
                {
                    if (error_MasterSub != "")
                        error_Exception = "Translation failed." + "\n" + "Validation error found while translating the following documents" + "\n" + " \"" + error_MasterSub + "\n" + error;
                    else
                        error_Exception = "Translated document is invalid." + "\n\n" + error;
                }
                else
                {
                    if (error_MasterSub != "")
                        error_Exception = "Translation failed." + "\n\n" + "Validation error found while translating the following documents" + "\n" + " \"" + error_MasterSub + "\n" + error;
                }

            }
            catch (UnauthorizedAccessException a)
            {
                xsd.Close();
                //dont have access permission
                String error = a.Message;
                report.AddLog("", "Validation Error of translated DAISY File: \n" + error, Report.ERROR_LEVEL);
            }
            catch (Exception a)
            {
                xsd.Close();
                //and other things that could go wrong
                String error = a.Message;
                report.AddLog("", "Validation Error of translated DAISY File: \n" + error, Report.ERROR_LEVEL);
            }
        }

        /* Function to Capture all the Validity Errors*/
        public void MyValidationEventHandler(object sender, ValidationEventArgs args)
        {
            isValid = false;
            error += " Line Number : " + args.Exception.LineNumber + " and " +
             " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
             " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;
        }

        /*Function to delete MathMl DTDs*/
        public void DeleteMath(String fileName)
        {
            DeleteFile(fileName + "\\mathml2.DTD");
            DeleteFile(fileName + "\\mathml2-qname-1.mod");
            Directory.Delete(fileName + "\\iso8879", true);
            Directory.Delete(fileName + "\\iso9573-13", true);
            Directory.Delete(fileName + "\\mathml", true);

        }

        /*Function to delete MathMl DTDs*/
        public void DeleteFile(String file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        /*Function to copy MATHML DTD to destination */
        public void CopyMATHToDestinationfolder(String outputFile)
        {
            string fileName = "";
            fileName = outputFile + "\\mathml2-qname-1.mod";
            CopyingAssemblyFile(fileName, "mathml2-qname-1.mod");
            fileName = outputFile + "\\mathml2.DTD";
            CopyingAssemblyFile(fileName, "mathml2.DTD");

            for (int i = 0; i < MathList8879.Count; i++)
            {
                Directory.CreateDirectory(outputFile + "\\iso8879");
                fileName = outputFile + "\\iso8879\\" + MathList8879[i].ToString();
                CopyingAssemblyFile(fileName, MathList8879[i].ToString());

            }

            for (int i = 0; i < MathList9573.Count; i++)
            {
                Directory.CreateDirectory(outputFile + "\\iso9573-13");
                fileName = outputFile + "\\iso9573-13\\" + MathList9573[i].ToString();
                CopyingAssemblyFile(fileName, MathList9573[i].ToString());
            }

            for (int i = 0; i < MathListmathml.Count; i++)
            {
                Directory.CreateDirectory(outputFile + "\\mathml");
                fileName = outputFile + "\\mathml\\" + MathListmathml[i].ToString();
                CopyingAssemblyFile(fileName, MathListmathml[i].ToString());

            }
        }

        /* Function which copies the Files to the destination folder*/
        public void CopyingAssemblyFile(String fileName, String indFilename)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;

            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(indFilename))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            StreamWriter writer = new StreamWriter(fileName);
            string data = reader.ReadToEnd();
            writer.Write(data);
            reader.Close();
            writer.Close();
        }

        /* Function which merge subdocument.xml to the master.xml*/
        public void MergeXml(string outputFile, XmlDocument mergeDoc, String rId, String inputFile)
        {
            try
            {
                XmlNode tempNode = null;
                XmlDocument tempDoc = new XmlDocument();
                tempDoc.Load(outputFile);

                tempDoc = SetFootnote(tempDoc, "subDoc" + subDocFootNum);

                for (int i = 0; i < tempDoc.SelectSingleNode("//head").ChildNodes.Count; i++)
                {
                    tempNode = tempDoc.SelectSingleNode("//head").ChildNodes[i];

                    if (tempNode.Attributes[0].Value == "dc:Language")
                    {
                        if (!mergeDocLanguage.Contains(tempNode.Attributes[1].Value))
                        {
                            mergeDocLanguage.Add(tempNode.Attributes[1].Value);
                        }
                    }
                }

                for (int i = 0; i < tempDoc.SelectSingleNode("//bodymatter").ChildNodes.Count; i++)
                {
                    tempNode = tempDoc.SelectSingleNode("//bodymatter").ChildNodes[i];

                    if (tempNode != null)
                    {
                        XmlNode addBodyNode = mergeDoc.ImportNode(tempNode, true);
                        if (addBodyNode != null)
                            mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']").ParentNode.InsertBefore(addBodyNode, mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']"));
                    }
                }

                tempNode = tempDoc.SelectSingleNode("//frontmatter/level1[@class='print_toc']");
                if (tempNode != null)
                {
                    if (!lostElements.Contains("TOC is not translated" + " for " + Path.GetFileName(inputFile)))
                    {
                        lostElements.Add("TOC is not translated" + " for " + Path.GetFileName(inputFile));
                    }

                }

                mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']").ParentNode.RemoveChild(mergeDoc.SelectSingleNode("//subdoc[@rId='" + rId + "']"));

                XmlNode node = tempDoc.SelectSingleNode("//rearmatter");

                if (node != null)
                {
                    for (int i = 0; i < tempDoc.SelectSingleNode("//rearmatter").ChildNodes.Count; i++)
                    {
                        tempNode = tempDoc.SelectSingleNode("//rearmatter").ChildNodes[i];

                        if (tempNode != null)
                        {
                            XmlNode addRearNode = mergeDoc.ImportNode(tempNode, true);
                            if (addRearNode != null)
                                mergeDoc.LastChild.LastChild.LastChild.AppendChild(addRearNode);
                        }
                    }
                }
                subDocFootNum++;
            }
            catch (Exception e)
            {
                error_MasterSub = error_MasterSub + "\n" + " \"" + inputFile + "\"";
                error_MasterSub = error_MasterSub + "\n" + "Validation error:" + "\n" + e.Message + "\n";
            }
        }

        /* Function which copies the DTD to the destination folder*/
        public void CopyDTDToDestinationfolder(String outputFile)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            string fileName = outputFile + "\\dtbook-2005-3.dtd";
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("dtbook-2005-3.dtd"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            writer.Write(data);
            writer.Close();

        }

        /* Function which creates language info of all sub documents in master.xml*/
        public void SetLanguage(XmlDocument mergeXmlDoc)
        {
            XmlNodeList languageList = mergeXmlDoc.SelectNodes("//meta[@name='dc:Language']");

            for (int i = 0; i < languageList.Count; i++)
            {
                if (mergeDocLanguage.Contains(languageList[i].Attributes[1].Value))
                {
                    int indx = mergeDocLanguage.IndexOf(languageList[i].Attributes[1].Value);
                    mergeDocLanguage.RemoveAt(indx);
                }
            }

            for (int i = 0; i < mergeDocLanguage.Count; i++)
            {
                XmlElement tempLang = mergeXmlDoc.CreateElement("meta");
                tempLang.SetAttribute("name", "dc:Language");
                tempLang.SetAttribute("content", mergeDocLanguage[i].ToString());
                mergeXmlDoc.SelectNodes("//head").Item(0).AppendChild(tempLang);
            }
        }

        /* Function which removes subdoc elements from the master.xml*/
        public void RemoveSubDoc(XmlDocument mergeXmlDoc)
        {
            XmlNodeList subDocList = mergeXmlDoc.SelectNodes("//subdoc");

            if (subDocList != null)
            {
                for (int i = 0; i < subDocList.Count; i++)
                {
                    subDocList.Item(i).ParentNode.RemoveChild(subDocList.Item(i));
                }
            }
        }

        /* Function which merges subdocument.xml and master.xml*/
        public void MergeXml(string outputFile, XmlDocument mergeDoc, String inputFile)
        {
            try
            {
                XmlNode tempNode = null;
                XmlDocument tempDoc = new XmlDocument();
                tempDoc.Load(outputFile);

                tempDoc = SetFootnote(tempDoc, "subDoc" + subDocFootNum);

                for (int i = 0; i < tempDoc.SelectSingleNode("//head").ChildNodes.Count; i++)
                {
                    tempNode = tempDoc.SelectSingleNode("//head").ChildNodes[i];

                    if (tempNode.Attributes[0].Value == "dc:Language")
                    {
                        if (!mergeDocLanguage.Contains(tempNode.Attributes[1].Value))
                        {
                            mergeDocLanguage.Add(tempNode.Attributes[1].Value);
                        }
                    }
                }


                for (int i = 0; i < tempDoc.SelectSingleNode("//bodymatter").ChildNodes.Count; i++)
                {
                    tempNode = tempDoc.SelectSingleNode("//bodymatter").ChildNodes[i];

                    if (tempNode != null)
                    {
                        XmlNode addBodyNode = mergeDoc.ImportNode(tempNode, true);
                        if (addBodyNode != null)
                            mergeDoc.LastChild.LastChild.FirstChild.NextSibling.AppendChild(addBodyNode);
                    }
                }

                tempNode = tempDoc.SelectSingleNode("//frontmatter/level1[@class='print_toc']");
                if (tempNode != null)
                {
                    if (!lostElements.Contains("TOC is not translated" + " for " + Path.GetFileName(inputFile)))
                    {
                        lostElements.Add("TOC is not translated" + " for " + Path.GetFileName(inputFile));
                    }

                }


                XmlNode node = tempDoc.SelectSingleNode("//rearmatter");

                if (node != null)
                {
                    for (int i = 0; i < tempDoc.SelectSingleNode("//rearmatter").ChildNodes.Count; i++)
                    {
                        tempNode = tempDoc.SelectSingleNode("//rearmatter").ChildNodes[i];

                        if (tempNode != null)
                        {
                            XmlNode addRearNode = mergeDoc.ImportNode(tempNode, true);
                            if (addRearNode != null)
                                mergeDoc.LastChild.LastChild.LastChild.AppendChild(addRearNode);
                        }
                    }
                }
                subDocFootNum++;
            }
            catch (Exception e)
            {
                error_MasterSub = error_MasterSub + "\n" + " \"" + inputFile + "\"";
            }
        }

        /* Function which deletes data and appends data of the xml*/
        public void ReplaceData(String fileName)
        {
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            if (!data.Contains("</mml:math>"))
            {
                data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd' >", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>");
                data = data.Replace("<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"", "<dtbook version=\"" + "2005-3\" xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"");
            }
            else
            {
                data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'[<!ENTITY % MATHML.prefixed \"INCLUDE\" ><!ENTITY % MATHML.prefix \"mml\"><!ENTITY % Schema.prefix \"sch\"><!ENTITY % XLINK.prefix \"xlp\"><!ENTITY % MATHML.Common.attrib \"xlink:href    CDATA       #IMPLIED xlink:type     CDATA       #IMPLIED   class          CDATA       #IMPLIED  style          CDATA       #IMPLIED  id             ID          #IMPLIED  xref           IDREF       #IMPLIED  other          CDATA       #IMPLIED   xmlns:dtbook   CDATA       #FIXED 'http://www.daisy.org/z3986/2005/dtbook/' dtbook:smilref CDATA       #IMPLIED\"><!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>%mathML2;<!ENTITY % externalFlow \"| mml:math\"><!ENTITY % externalNamespaces \"xmlns:mml CDATA #FIXED 'http://www.w3.org/1998/Math/MathML'\">]>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>");
                cutData = data.Substring(95, 1091);
                data = data.Remove(95, 1091);
                data = data.Replace("<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"", "<dtbook version=\"" + "2005-3\"");
            }
            writer.Write(data);
            writer.Close();
        }

        /* Function which merges subdocument.xml and master.xml*/
        public void ReplaceData(String fileName, bool value)
        {
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            if (value)
            {
                if (!data.Contains("</mml:math>"))
                {
                    data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>");
                    data = data.Replace("<dtbook version=\"" + "2005-3\" xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xml:lang=", "<dtbook version=\"" + "2005-3\" xml:lang=");
                }
                else
                {
                    tempData = cutData.Replace("<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'", "<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'");
                    tempData = tempData.Replace("<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">", "<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>");
                    data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>" + tempData);
                }
            }
            else
            {
                if (!data.Contains("</mml:math>"))
                {
                    data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'>");
                    data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                }
                else
                {
                    data = data.Replace(tempData, cutData);
                    data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                }
            }
            writer.Write(data);
            writer.Close();
        }

        /* Function which creates unique ID to page numbers*/
        public void SetPageNum(XmlDocument mergeXmlDoc)
        {
            XmlNodeList pageList = mergeXmlDoc.SelectNodes("//pagenum");
            for (int i = 1; i <= pageList.Count; i++)
            {
                mergeXmlDoc.SelectNodes("//pagenum").Item(i - 1).Attributes[1].InnerText = "page" + i.ToString();
            }
        }

        /* Function which creates unique ID to Images*/
        public void SetImage(XmlDocument mergeXmlDoc)
        {
            XmlNodeList imageList = mergeXmlDoc.SelectNodes("//img");
            int j = 0;
            for (int i = 1; i <= imageList.Count; i++)
            {
                if (mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText.StartsWith("rId"))
                {
                    mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText = "rId" + j.ToString();
                    j++;
                }
            }
            SetCaption_Image(mergeXmlDoc);
        }

        /* Function which creates unique ID to the Image Captions */
        public void SetCaption_Image(XmlDocument mergeXmlDoc)
        {
            XmlNodeList captionList = mergeXmlDoc.SelectNodes("//caption");
            for (int i = 1; i <= captionList.Count; i++)
            {
                XmlNode prevNode = mergeXmlDoc.SelectNodes("//caption").Item(i - 1).PreviousSibling;
                if (prevNode != null)
                {
                    String rId = prevNode.Attributes[0].InnerText;
                    mergeXmlDoc.SelectNodes("//caption").Item(i - 1).Attributes[0].InnerText = rId;
                }
            }
        }

        /* Function which creates unique ID to the Footnotes */
        public XmlDocument SetFootnote(XmlDocument mergeXmlDoc, String SubDocFootnum)
        {
            int footnoteCount = 1, endnoteCount = 1;
            XmlNodeList noteList = mergeXmlDoc.SelectNodes("//note");
            if (noteList != null)
            {

                for (int i = 1; i <= noteList.Count; i++)
                {
                    if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Footnote")
                    {
                        mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "footnote-" + footnoteCount.ToString();
                        footnoteCount++;
                    }
                    if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Endnote")
                    {
                        mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "endnote-" + endnoteCount.ToString();
                        endnoteCount++;
                    }
                }
            }

            footnoteCount = 1;
            endnoteCount = 1;
            noteList = mergeXmlDoc.SelectNodes("//noteref");
            if (noteList != null)
            {

                for (int i = 1; i <= noteList.Count; i++)
                {
                    if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Footnote")
                    {
                        mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "footnote-" + footnoteCount.ToString();
                        footnoteCount++;
                    }
                    if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Endnote")
                    {
                        mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "endnote-" + endnoteCount.ToString();
                        endnoteCount++;
                    }
                }
            }

            return mergeXmlDoc;

        }

        #endregion


    }
}
