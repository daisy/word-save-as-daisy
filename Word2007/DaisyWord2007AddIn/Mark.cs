using System;
using System.Xml;
using System.Data;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Collections;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.ComponentModel;
using MSword = Microsoft.Office.Interop.Word;
using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY;

namespace Daisy.SaveAsDAISY.Addins.Word2007 {
    public partial class Mark : Form {
        bool checkAbbrAcr;
        int checkCustomXml = 0;
        MSword.Document currentDoc;
        public AddinResources addinLib;
        String pronounceAbbrAcr = "No", allOccurences = "No";
        XmlDocument customXml = new XmlDocument();
        XmlDocument manageAbbrAcr = new XmlDocument();

        /// <summary>
        /// Constructor which loads different UI for Abbreviations/Acronyms
        /// </summary>
        /// <param name="doc">Word Document</param>
        /// <param name="value">Whether call is from Abbreviation Button or Acronym Button</param>
        public Mark(MSword.Document doc, bool valueAbbrAcr) {
            InitializeComponent();
            currentDoc = doc;
            checkAbbrAcr = valueAbbrAcr;
            this.addinLib = new Daisy.SaveAsDAISY.AddinResources();

            //Checking whether call is from Abbreviation button or Acronym button
            //to diifferenciate the Look and Feel of UI
            //If true,it is Abbreviation Button
            if (checkAbbrAcr) {
                this.Text = addinLib.GetString("MarkAbbr");
                this.label1.Text = "&Full form of \"" + currentDoc.Application.Selection.Text.Trim() + "\" :";
                this.groupBox_Settings.Height = 45;
                this.groupBox_Settings.Width = 256;
            }
            //Otherwise it is Acronym Button
            else {
                this.Text = addinLib.GetString("MarkAcr");
                this.label1.Text = "&Full form of \"" + currentDoc.Application.Selection.Text.Trim() + "\" :";
                this.cBx_Pronounce.Visible = true;
            }
        }

        /// <summary>
        /// Function to set the value of pronounce Attribute for Abbreviation/Acronyms
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cBx_Pronounce_CheckedChanged(object sender, EventArgs e) {
            if (cBx_Pronounce.Checked)
                pronounceAbbrAcr = "Yes";
            else
                pronounceAbbrAcr = "No";
        }

        /// <summary>
        /// Function to set the value whether to apply Abbreviations\Acronyms
        /// for all the occurences of a particular Word in the current document.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cBx_AllOccurences_CheckedChanged(object sender, EventArgs e) {
            if (cBx_AllOccurences.Checked)
                allOccurences = "Yes";
            else
                allOccurences = "No";
        }


        /// <summary>
        /// Function to check whether a Bookmark is an Abbreviation or Acronym
        /// </summary>
        /// <param name="rngDoc">Text in the Word document</param>
        /// <returns>String saying whether selected text is already an Abbreviaton\Acronym</returns>
        public String ValidateBookMark(MSword.Range rngDoc) {
            String bookMarkExists = "Nobookmarks";
            if (rngDoc.Bookmarks.Count > 0) {
                foreach (object item in currentDoc.Bookmarks) {
                    if (((MSword.Bookmark)item).Range.Text.Trim() == rngDoc.Text.Trim()) {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
                            bookMarkExists = "Abbrtrue";
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                            bookMarkExists = "Acrtrue";
                    }
                }
            }
            return bookMarkExists;
        }

        /// <summary>
        /// Function which marks the selected text as an Abbreviation\Acronym
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Mark_Click(object sender, EventArgs e) {
            try {
                if (currentDoc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection) {
                    CustomXMLParts xmlParts = currentDoc.CustomXMLParts;
                    //Checking whether CustomXML is there or not in the Current Document
                    for (int i = 1; i <= xmlParts.Count; i++) {
                        if (xmlParts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            checkCustomXml = 1;
                        }
                    }

                    //If CustomXML is not there in the Current Document
                    if (checkCustomXml == 0) {
                        //Creating Custom XML document for the Current Document
                        XmlElement elmtManage = manageAbbrAcr.CreateElement("Manage");
                        manageAbbrAcr.AppendChild(elmtManage);
                        elmtManage.SetAttribute("xmlns", "http://Daisy-OpenXML/customxml");

                        XmlElement elmtAbbreviations = manageAbbrAcr.CreateElement("Abbreviations");
                        elmtManage.AppendChild(elmtAbbreviations);

                        XmlElement elmtAcronyms = manageAbbrAcr.CreateElement("Acronyms");
                        elmtManage.AppendChild(elmtAcronyms);

                        //If user is not selected "allOccurrences" checkbox 
                        if (allOccurences == "No") {
                            //If call is from Abbreviation button
                            if (checkAbbrAcr) {
                                object val = System.Reflection.Missing.Value;
                                String nameAbbr = "Abbreviations" + GenerateId().ToString();
                                //Adding Bookmark for the selected Text
                                currentDoc.Application.Selection.Bookmarks.Add(nameAbbr, ref val);

                                //Updating the CustomXML
                                XmlElement elmtItem = manageAbbrAcr.CreateElement("Item");
                                elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                                elmtItem.SetAttribute("FullAbbr", tBx_MarkFullForm.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());
                                elmtAbbreviations.AppendChild(elmtItem);

                                //Adding the CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(manageAbbrAcr.InnerXml, System.Reflection.Missing.Value);
                                this.Close();
                            }
                            //If call is from Acronym button
                            else {
                                object val = System.Reflection.Missing.Value;
                                String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();
                                //Adding Bookmark for the selected Text
                                currentDoc.Application.Selection.Bookmarks.Add(nameAcr, ref val);

                                //Updating the CustomXML
                                XmlElement elmtItem = manageAbbrAcr.CreateElement("Item");
                                elmtItem.SetAttribute("AcronymName", nameAcr);
                                elmtItem.SetAttribute("FullAcr", tBx_MarkFullForm.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());
                                elmtAcronyms.AppendChild(elmtItem);

                                //Adding the CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(manageAbbrAcr.InnerXml, System.Reflection.Missing.Value);
                                this.Close();
                            }
                        }
                        //If user is selected "allOccurrences" checkbox 
                        else {
                            //If call is from Abbreviation button
                            if (checkAbbrAcr) {
                                object missing = System.Type.Missing;
                                Object startDoc = 0;
                                Object endDoc = currentDoc.Characters.Count;
                                MSword.Range rngDoc = currentDoc.Range(ref startDoc, ref endDoc);
                                MSword.Find fndDoc = rngDoc.Find;

                                fndDoc.ClearFormatting();
                                fndDoc.Forward = true;
                                fndDoc.Text = currentDoc.Application.Selection.Text.Trim();
                                ExecuteFind(fndDoc);

                                //Applying Abbreviation for all the occurrences of the selected Word
                                while (fndDoc.Found) {
                                    if (ValidateBookMark(rngDoc) == "Nobookmarks") {
                                        String nameAbbr = "Abbreviations" + GenerateId().ToString();
                                        rngDoc.Bookmarks.Add(nameAbbr, ref missing);


                                        XmlElement elmtItem = manageAbbrAcr.CreateElement("Item");
                                        elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                                        elmtItem.SetAttribute("FullAbbr", tBx_MarkFullForm.Text.TrimEnd());
                                        elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                                        elmtAbbreviations.AppendChild(elmtItem);
                                    }
                                    ExecuteFind(fndDoc);
                                }

                                //Adding the CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(manageAbbrAcr.InnerXml, System.Reflection.Missing.Value);
                                this.Close();
                            }
                            //If call is from Acronym button
                            else {
                                object missing = System.Type.Missing;
                                Object startDoc = 0;
                                Object endDoc = currentDoc.Characters.Count;
                                MSword.Range rngDoc = currentDoc.Range(ref startDoc, ref endDoc);
                                MSword.Find fndDoc = rngDoc.Find;

                                fndDoc.ClearFormatting();
                                fndDoc.Forward = true;
                                fndDoc.Text = currentDoc.Application.Selection.Text.Trim();
                                ExecuteFind(fndDoc);

                                //Applying Abbreviation for all the occurrences of the selected Word
                                while (fndDoc.Found) {
                                    if (ValidateBookMark(rngDoc) == "Nobookmarks") {
                                        String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();
                                        rngDoc.Bookmarks.Add(nameAcr, ref missing);

                                        XmlElement elmtItem = manageAbbrAcr.CreateElement("Item");
                                        elmtItem.SetAttribute("AcronymName", nameAcr);
                                        elmtItem.SetAttribute("FullAcr", tBx_MarkFullForm.Text.TrimEnd());
                                        elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                                        elmtAcronyms.AppendChild(elmtItem);
                                    }
                                    ExecuteFind(fndDoc);
                                }

                                //Adding the CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(manageAbbrAcr.InnerXml, System.Reflection.Missing.Value);
                                this.Close();
                            }
                        }
                    }
                    //If already CustomXML is there in the Current Document
                    else {
                        //If user is not selected "allOccurrences" checkbox 
                        if (allOccurences == "No") {
                            Individual();
                        } else {
                            AllOccurences();
                        }
                    }
                } else {
                    MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Function which Marks all occurences of a particular word as Abbreviations or Acronyms
        /// </summary>
        public void AllOccurences() {
            CustomXMLParts xmlparts = currentDoc.CustomXMLParts;

            //If call is from Abbreviation button
            if (checkAbbrAcr) {
                object missing = System.Type.Missing;
                Object startDoc = 0;
                Object endDoc = currentDoc.Characters.Count;
                MSword.Range rngDoc = currentDoc.Range(ref startDoc, ref endDoc);
                MSword.Find fndDoc = rngDoc.Find;

                fndDoc.ClearFormatting();
                fndDoc.Forward = true;
                fndDoc.Text = currentDoc.Application.Selection.Text.Trim();
                ExecuteFind(fndDoc);

                //Applying Abbreviation for all the occurrences of the selected Word
                while (fndDoc.Found) {
                    String nameAbbr = "Abbreviations" + GenerateId().ToString();

                    if (ValidateBookMark(rngDoc) == "Nobookmarks") {
                        rngDoc.Bookmarks.Add(nameAbbr, ref missing);

                        for (int i = 1; i <= xmlparts.Count; i++) {
                            if (xmlparts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                                String partxml = xmlparts[i].XML;
                                customXml.LoadXml(partxml);

                                XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                                elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                                elmtItem.SetAttribute("FullAbbr", tBx_MarkFullForm.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                                customXml.FirstChild.FirstChild.AppendChild(elmtItem);

                                //Deleting the Current CustomXML
                                xmlparts[i].Delete();

                                //Adding the updated CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(customXml.InnerXml, System.Reflection.Missing.Value);
                            }
                        }
                    }
                    ExecuteFind(fndDoc);
                }
                this.Close();
            } else {
                object missing = System.Type.Missing;
                Object start = 0;
                Object end = currentDoc.Characters.Count;
                MSword.Range rngDoc = currentDoc.Range(ref start, ref end);
                MSword.Find fnd = rngDoc.Find;

                fnd.ClearFormatting();
                fnd.Forward = true;
                fnd.Text = currentDoc.Application.Selection.Text.Trim();
                ExecuteFind(fnd);
                while (fnd.Found) {
                    String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();

                    if (ValidateBookMark(rngDoc) == "Nobookmarks") {
                        rngDoc.Bookmarks.Add(nameAcr, ref missing);


                        for (int i = 1; i <= xmlparts.Count; i++) {
                            if (xmlparts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                                String partxml = xmlparts[i].XML;
                                customXml.LoadXml(partxml);

                                XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                                elmtItem.SetAttribute("AcronymName", nameAcr);
                                elmtItem.SetAttribute("FullAcr", tBx_MarkFullForm.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                                customXml.FirstChild.FirstChild.NextSibling.AppendChild(elmtItem);

                                //Deleting the Current CustomXML
                                xmlparts[i].Delete();

                                //Adding the updated CustomXml to the Current Document
                                CustomXMLPart part = currentDoc.CustomXMLParts.Add(customXml.InnerXml, System.Reflection.Missing.Value);
                            }
                        }
                    }
                    ExecuteFind(fnd);
                }
                this.Close();
            }
        }

        /// <summary>
        /// Function which Marks a individual word as Abbreviation or Acronym
        /// </summary>
        public void Individual() {
            CustomXMLParts xmlparts = currentDoc.CustomXMLParts;

            //If call is from Abbreviation button
            if (checkAbbrAcr) {
                object val = System.Reflection.Missing.Value;
                String nameAbbr = "Abbreviations" + GenerateId().ToString();
                currentDoc.Application.Selection.Bookmarks.Add(nameAbbr, ref val);

                for (int i = 1; i <= xmlparts.Count; i++) {
                    if (xmlparts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[i].XML;
                        customXml.LoadXml(partxml);

                        XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                        elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                        elmtItem.SetAttribute("FullAbbr", tBx_MarkFullForm.Text.TrimEnd());
                        elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                        customXml.FirstChild.FirstChild.AppendChild(elmtItem);

                        //Deleting the Current CustomXML
                        xmlparts[i].Delete();

                        //Adding the updated CustomXml to the Current Document
                        CustomXMLPart part = currentDoc.CustomXMLParts.Add(customXml.InnerXml, System.Reflection.Missing.Value);
                        this.Close();
                    }
                }
            } else {
                object val = System.Reflection.Missing.Value;
                String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();
                currentDoc.Application.Selection.Bookmarks.Add(nameAcr, ref val);

                for (int i = 1; i <= xmlparts.Count; i++) {

                    if (xmlparts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[i].XML;
                        customXml.LoadXml(partxml);

                        XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                        elmtItem.SetAttribute("AcronymName", nameAcr);
                        elmtItem.SetAttribute("FullAcr", tBx_MarkFullForm.Text.TrimEnd());
                        elmtItem.SetAttribute("OriginalText", currentDoc.Application.Selection.Text.Trim());

                        customXml.FirstChild.FirstChild.NextSibling.AppendChild(elmtItem);

                        //Deleting the Current CustomXML
                        xmlparts[i].Delete();

                        //Adding the updated CustomXml to the Current Document
                        CustomXMLPart part = currentDoc.CustomXMLParts.Add(customXml.InnerXml, System.Reflection.Missing.Value);
                        this.Close();
                    }
                }
            }
        }

        /*Function which generates a random ID*/
        public long GenerateId() {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        /*Function which finds all occurences of a particular word*/
        public Boolean ExecuteFind(MSword.Find find) {
            object missing = System.Type.Missing;
            object wholeword = true;
            object MatchCase = true;

            return find.Execute(ref missing, ref MatchCase, ref wholeword,
              ref missing, ref missing, ref missing,
              ref missing, ref missing, ref missing, ref missing, ref missing,
              ref missing, ref missing, ref missing,
              ref missing);
        }

        private void btn_Cancel_Click(object sender, EventArgs e) {
            this.Close();
        }

    }
}