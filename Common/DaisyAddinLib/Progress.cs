using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Resources;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Windows.Forms;
using System.IO.Packaging;
using Daisy.DaisyConverter.DaisyConverterLib;

namespace Daisy.DaisyConverter.DaisyConverterLib
{
    public partial class Progress : Form
    {
        private AbstractConverter converter;
        XmlDocument mergeXmlDoc;
        ArrayList mergeDocLanguage, subList;
        private Exception exception;
        Hashtable table;
        private int size;
        String cutData;
        String tempData = "";
        int subDocFootNum;
        private static bool isValid;
        private bool cancel, converting, computeSize;
        private static string error;
        private ResourceManager manager;
        private static string error_MasterSub = "";
        ArrayList MathList8879, MathList9573, MathListmathml;
        String tempInputFile, outputfilepath, category = "", error_Exception = "", output_Pipeline;
        ArrayList lostElements = new ArrayList();
        String docName = "";
        string control = "";
        String errorText = "";
        string path_For_Pipeline = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms";
        private Hashtable listMathMl;
        private Hashtable multipleMathMl;
        /*Constructor*/
        public Progress(AbstractConverter converter, ResourceManager manager, String outputfile, ArrayList list, Hashtable table, String inputFile, string control, Hashtable listMathMl, string output_Pipeline)
        {
            InitializeComponent();
            this.converter = converter;
            this.manager = manager;
            this.outputfilepath = outputfile;
            subList = list;
            this.table = table;
            this.progressBar1.Maximum = subList.Count;
            this.tempInputFile = inputFile;
            subDocFootNum = 1;
            lblProgress.Text = "Loading translation details. Please wait for a while...";
            AddMathmlDtds();
            this.control = control;
            this.listMathMl = listMathMl;
            this.output_Pipeline = output_Pipeline;
        }

        /* Bunch of Docs Constructor*/
        public Progress(AbstractConverter converter, ResourceManager manager, String outputfile, ArrayList list, String category, Hashtable propTable, string control, Hashtable multipleMathMl, string output_Pipeline)
        {
            InitializeComponent();
            this.converter = converter;
            this.manager = manager;
            this.outputfilepath = outputfile;
            subList = list;
            this.table = propTable;
            this.progressBar1.Maximum = subList.Count;
            this.category = category;
            subDocFootNum = 1;
            lblProgress.Text = "Loading translation details. Please wait for a while...";
            AddMathmlDtds();
            this.control = control;
            this.multipleMathMl = multipleMathMl;
            this.output_Pipeline = output_Pipeline;
        }

        /* Returns all errors occured during translation*/
        public ArrayList LostElements
        {
            get
            {
                return lostElements;
            }
        }

        /* Returns boolean value if there are any errors occured during translation*/
        public bool HasLostElements
        {
            get
            {
                return lostElements.Count > 0;
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

        /*Function to Load the form*/
        private void Progress_Load(object sender, EventArgs e)
        {
        }

        /*Function which starts the Translation process*/
        private void Progress_Activated(object sender, EventArgs e)
        {
            if (!converting)
            {
                converting = true;
                Application.DoEvents();
                if (category == "individual")
                     DoTranslation_Ind();
                else
                    DoTranslation();
                converting = false;
            }
        }

        /* Function which Increments the value of progressbar*/
        private void ProgressMessageInterceptorMaster(object sender, EventArgs e)
        {
            if (this.computeSize)
            {
                this.size++;
            }
            else
            {
                this.progressBar1.Increment(1);
                if (Directory.Exists(path_For_Pipeline))
                {
                    if ((control != "DaisySingle") && (control != "DaisyTabSingle") && (control != "DaisyMultiple") && (control != "DaisyTabMultiple") && (control != "Button2"))
                    {
                        float percentShow = (this.progressBar1.Value * 100) / this.progressBar1.Maximum;
                        lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                        lblProgress.Visible = true;
                        if (percentShow == 100.00)
                        {
							lblProgress.Text = "Initiating daisy pipeline lite.Please wait...";
                            lblProgress.Visible = true;
                        }
                    }
                    else
                    {
                        float percentShow = (this.progressBar1.Value * 100) / this.progressBar1.Maximum;
                        lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                        lblProgress.Visible = true;
                    }
                }
                else
                {
                    float percentShow = (this.progressBar1.Value * 100) / this.progressBar1.Maximum;
                    lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                    lblProgress.Visible = true;
                }
            }
            Application.DoEvents();
            if (cancel)
            {
                throw new CancelledException();
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
                    messageValue = manager.GetString(messageKey[0].Substring(0, index));

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
                    messageValue = manager.GetString(messageKey[0]);
                }

                if (messageValue != null && !lostElements.Contains(messageKey[1] + messageValue + " for " + Path.GetFileName(docName)))
                {
                    lostElements.Add(messageKey[1] + messageValue + " for " + Path.GetFileName(docName));
                }
            }
            else
            {
                int index = message.IndexOf('%');
                // parameters substitution
                if (index > 0)
                {
                    string[] param = message.Substring(index + 1).Split(new char[] { '%' });
                    messageValue = manager.GetString(message.Substring(0, index));

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
                    messageValue = manager.GetString(message);
                }

                if (messageValue != null && !lostElements.Contains(messageValue + " for " + Path.GetFileName(docName)))
                {
                    lostElements.Add(messageValue + " for " + Path.GetFileName(docName));
                }
            }


        }


        #region Multiple OOXML functions

        /* Function which Translates the current document along with its sub documents*/
        public void DoTranslation()
        {
            try
            {
                error_MasterSub = "";
                mergeXmlDoc = new XmlDocument();
                mergeDocLanguage = new ArrayList();
                this.computeSize = true;
                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    docName = splt[0];
                    converter.RemoveMessageListeners();
                    converter.AddProgressMessageListenerMaster(new AbstractConverter.MessageListener(ProgressMessageInterceptorMaster));
                    converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
					converter.Transform(splt[0], null, table, null, true, "");
                }
                this.progressBar1.Maximum = this.size;
                this.computeSize = false;

                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    String outputFile = outputfilepath + "\\" + Path.GetFileNameWithoutExtension(splt[0]) + ".xml";
                    String ridOutputFile = splt[1];
                    docName = splt[0];
                    this.converter.Transform(splt[0], outputFile, table, (Hashtable)listMathMl["Doc" + i], true, output_Pipeline);
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
                mergeXmlDoc.Save(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
                ReplaceData(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml", true);
                CopyDTDToDestinationfolder(outputfilepath);
                CopyMATHToDestinationfolder(outputfilepath);
                XmlValidation(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
                ReplaceData(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml", false);
                if (File.Exists(outputfilepath + "\\dtbook-2005-3.dtd"))
                {
                    File.Delete(outputfilepath + "\\dtbook-2005-3.dtd");
                }
                DeleteMath(outputfilepath, true);
                WorkComplete(null);
            }
            catch (Exception e)
            {
                WorkComplete(e);
                error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + error_MasterSub + "\n" + "Problem is:" + "\n" + e.Message + "\n";

            }
        }

        /* Function which Translates a set of sub documents*/
        public void DoTranslation_Ind()
        {
            try
            {
                error_MasterSub = "";
                mergeXmlDoc = new XmlDocument();
                mergeDocLanguage = new ArrayList();

                this.computeSize = true;
                for (int i = 0; i < subList.Count; i++)
                {
                    string[] splt = subList[i].ToString().Split('|');
                    docName = splt[0];
                    converter.RemoveMessageListeners();
                    converter.AddProgressMessageListenerMaster(new AbstractConverter.MessageListener(ProgressMessageInterceptorMaster));
                    converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
					converter.Transform(splt[0], null, table, null, true, "");
                }

                this.progressBar1.Maximum = this.size;
                this.computeSize = false;

                for (int i = 0; i < subList.Count; i++)
                {
                    String outputFile = Path.GetDirectoryName(outputfilepath) + "\\" + Path.GetFileNameWithoutExtension(subList[i].ToString()) + ".xml";
                    docName = subList[i].ToString();
                    this.converter.Transform(subList[i].ToString(), outputFile, table, (Hashtable)multipleMathMl["Doc" + i], true, output_Pipeline);
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

                SetPageNum(mergeXmlDoc);
                SetImage(mergeXmlDoc);
                SetLanguage(mergeXmlDoc);
                RemoveSubDoc(mergeXmlDoc);
                mergeXmlDoc.Save(outputfilepath);
                ReplaceData(outputfilepath, true);
                CopyDTDToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                CopyMATHToDestinationfolder(Path.GetDirectoryName(outputfilepath));
                XmlValidation(outputfilepath);
                ReplaceData(outputfilepath, false);
                if (File.Exists(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd"))
                {
                    File.Delete(Path.GetDirectoryName(outputfilepath) + "\\dtbook-2005-3.dtd");
                }
                DeleteMath(Path.GetDirectoryName(outputfilepath), true);
                WorkComplete(null);
            }
            catch (Exception e)
            {
                WorkComplete(e);
                error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(subList[0].ToString()) + "\"" + "\n" + "Validation error:" + "\n" + e.Message + "\n";
            }
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
                error_MasterSub = error_MasterSub + "\n" + "Validation error:" + "\n" + e.Message + "\n";
            }
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

        /* Function which validated the master.xml*/
        public void XmlValidation(String outFile)
        {
            isValid = true;
            error = "";
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

                XslCompiledTransform trans = new XslCompiledTransform(true);
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
                        error_Exception = manager.GetString("TranslationFailed") + "\n" + manager.GetString("WellDaisyFormat") + error_MasterSub + "\n" + error;
                    else
                        error_Exception = manager.GetString("InvalidMaster") + "\n\n" + error;
                }
                else
                {
                    if (error_MasterSub != "")
                        error_Exception = manager.GetString("TranslationFailed") + "\n\n" + manager.GetString("WellDaisyFormat") + error_MasterSub + "\n" + error;
                }

            }
            catch (UnauthorizedAccessException a)
            {
                xsd.Close();
                //dont have access permission
                error = a.Message;

                System.Windows.Forms.MessageBox.Show(error, "Daisy Translator", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception a)
            {
                xsd.Close();
                //and other things that could go wrong
                error = a.Message;

                System.Windows.Forms.MessageBox.Show(error, "Daisy Translator", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /* Function which returns the error messages*/
        public void MyValidationEventHandler(object sender, ValidationEventArgs args)
        {
            isValid = false;
            error += " Line Number : " + args.Exception.LineNumber + " and " +
             " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
             " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;
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


        #endregion

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

        /*Function which gives value according to exceptions*/
        private void WorkComplete(Exception e)
        {
            if (e == null)
            {
                DialogResult = DialogResult.OK;
            }
            else
            {
                if (e is CancelledException)
                {
                    DialogResult = DialogResult.Cancel;
                }
                else
                {
                    this.exception = e;
                    DialogResult = DialogResult.Abort;
                    string[] files = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY\");
                    foreach (string file in files)
                    {
                        if (file.Contains(".jpg") || file.Contains(".JPG") || file.Contains(".PNG") || file.Contains(".png"))
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
        }

        /* Returns the cancelled exception*/
        public class CancelledException : Exception
        {
        }

        /* Function Returns the exception*/
        public Exception Exception
        {
            get
            {
                return this.exception;
            }
        }

        #region MathML functions

        /*Function to create a list for all the required Mathml files*/
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


        /*Function to copy MathML files to Output folder*/
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

        /*Function which deletes the file*/
        public void DeleteMath(String fileName, bool value)
        {
            if (value)
            {
                DeleteFile(fileName + "\\mathml2.DTD");
                DeleteFile(fileName + "\\mathml2-qname-1.mod");
                Directory.Delete(fileName + "\\iso8879", true);
                Directory.Delete(fileName + "\\iso9573-13", true);
                Directory.Delete(fileName + "\\mathml", true);
            }
        }

        /*Function which deletes the file*/
        public void DeleteFile(String file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        #endregion

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(manager.GetString("QuitConfirmation"), "DAISY Translator", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                cancel = true;

            }
            else
                cancel = false;
        }

    }
}