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

namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    public partial class Mark2003 : Form
    {
        bool checkAbbrAcr;
        MSword.Document currentDoc;
        String pronounceAbbrAcr = "No", allOccurences = "No";
        XmlDocument dc = new XmlDocument();
        XmlDocument manageAbbrAcr = new XmlDocument();
        public AddinResources addinLib;
        object type = MsoDocProperties.msoPropertyTypeString;

        /// <summary>
        /// Constructor which loads different UI for Abbreviations/Acronyms
        /// </summary>
        /// <param name="doc">Word Document</param>
        /// <param name="value">Whether call is from Abbreviation Button or Acronym Button</param>
        public Mark2003(MSword.Document doc, bool valueAbbrAcr)
        {
            InitializeComponent();
            currentDoc = doc;
            checkAbbrAcr = valueAbbrAcr;
            this.addinLib = new Daisy.SaveAsDAISY.AddinResources();

            //Checking whether call is from Abbreviation button or Acronym button
            //to diifferenciate the Look and Feel of UI
            //If true,it is Abbreviation Button
            if (checkAbbrAcr)
            {
                this.Text = addinLib.GetString("MarkAbbr");
                this.label1.Text = "&Full form of \"" + currentDoc.Application.Selection.Text.Trim() + "\" :";

                this.groupBox_Settings.Height = 45;
                this.groupBox_Settings.Width = 256;
            }
            //Otherwise it is Acronym Button
            else
            {
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
        private void cBx_Pronounce_CheckedChanged(object sender, EventArgs e)
        {
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
        private void cBx_AllOccurences_CheckedChanged(object sender, EventArgs e)
        {
            if (cBx_AllOccurences.Checked)
                allOccurences = "Yes";
            else
                allOccurences = "No";
        }

        /// <summary>
        /// Function which Marks all occurences of a particular word as Abbreviations or Acronyms
        /// </summary>
        public void AllOccurences()
        {
            DocumentProperties xmlparts = (DocumentProperties)currentDoc.CustomDocumentProperties;

            //If call is from Abbreviation button
            if (checkAbbrAcr)
            {
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
                while (fndDoc.Found)
                {
                    if (ValidateBookMark(rngDoc) == "Nobookmarks")
                    {
                        string nameAbbr = "Abbreviations" + GenerateId().ToString();
                        //Adding the Bookamrk for the selected Word
                        rngDoc.Bookmarks.Add(nameAbbr, ref missing);

                        object value = currentDoc.Application.Selection.Text.Trim() + "$#$" + tBx_MarkFullForm.Text.TrimEnd();
                        //Adding a custom property to the current document
                        xmlparts.Add(nameAbbr, false, type, value, System.Reflection.Missing.Value);
                    }
                    ExecuteFind(fndDoc);
                }
                this.Close();
            }
            else
            {
                object missing = System.Type.Missing;
                Object startDoc = 0;
                Object endDoc = currentDoc.Characters.Count;
                MSword.Range rngDoc = currentDoc.Range(ref startDoc, ref endDoc);
                MSword.Find fndDoc = rngDoc.Find;

                fndDoc.ClearFormatting();
                fndDoc.Forward = true;
                fndDoc.Text = currentDoc.Application.Selection.Text.Trim();
                ExecuteFind(fndDoc);

                //Applying Acronymns for all the occurrences of the selected Word
                while (fndDoc.Found)
                {
                    if (ValidateBookMark(rngDoc) == "Nobookmarks")
                    {
                        String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();
                        //Adding the Bookamrk for the selected Word
                        rngDoc.Bookmarks.Add(nameAcr, ref missing);

                        object value = currentDoc.Application.Selection.Text.Trim() + "$#$" + tBx_MarkFullForm.Text.TrimEnd();
                        //Adding a custom property to the current document
                        xmlparts.Add(nameAcr, false, type, value, System.Reflection.Missing.Value);
                    }
                    ExecuteFind(fndDoc);
                }
                this.Close();
            }
        }

        /// <summary>
        /// Function which Marks a individual word as Abbreviation or Acronym
        /// </summary>
        public void Individual()
        {
            DocumentProperties xmlparts = (DocumentProperties)currentDoc.CustomDocumentProperties;

            //If call is from Abbreviation button
            if (checkAbbrAcr)
            {
                object missing = System.Type.Missing;
                object val = System.Reflection.Missing.Value;
                String nameAbbr = "Abbreviations" + GenerateId().ToString();
                //Adding bookmark for the selected Text
                currentDoc.Application.Selection.Bookmarks.Add(nameAbbr, ref val);

                object value = currentDoc.Application.Selection.Text.Trim() + "$#$" + tBx_MarkFullForm.Text.TrimEnd();
                //Adding a custom property to the current document
                xmlparts.Add(nameAbbr, false, type, value, System.Reflection.Missing.Value);
                this.Close();
            }
            else
            {
                object missing = System.Type.Missing;
                object val = System.Reflection.Missing.Value;
                String nameAcr = "Acronyms" + pronounceAbbrAcr + GenerateId().ToString();
                //Adding bookmark for the selected Text
                currentDoc.Application.Selection.Bookmarks.Add(nameAcr, ref val);

                object value = currentDoc.Application.Selection.Text.Trim() + "$#$" + tBx_MarkFullForm.Text.TrimEnd();
                //Adding a custom property to the current document
                xmlparts.Add(nameAcr, false, type, value, System.Reflection.Missing.Value);
                this.Close();
            }
        }
        /*Function which generates a random ID*/
        public long GenerateId()
        {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        /*Function which finds all occurences of a particular word*/
        public Boolean ExecuteFind(MSword.Find find)
        {
            object missing = System.Type.Missing;
            object wholeword = true;
            object MatchCase = true;

            return find.Execute(ref missing, ref MatchCase, ref wholeword,
              ref missing, ref missing, ref missing,
              ref missing, ref missing, ref missing, ref missing, ref missing,
              ref missing, ref missing, ref missing,
              ref missing);
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            //pdoc.Save();
            this.Close();
        }

        private void btn_Mark_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentDoc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                {
                    if (allOccurences == "Yes")
                        AllOccurences();
                    else
                        Individual();
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
        /// Function to check whether a Bookmark is an Abbreviation or Acronym
        /// </summary>
        /// <param name="rngDoc">Text in the Word document</param>
        /// <returns>String saying whether selected text is already an Abbreviaton\Acronym</returns>
        public String ValidateBookMark(MSword.Range rngDoc)
        {
            String bookMarkExists = "Nobookmarks";

            if (rngDoc.Bookmarks.Count > 0)
            {
                foreach (object item in currentDoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Range.Text.Trim() == rngDoc.Text.Trim())
                    {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
                            bookMarkExists = "Abbrtrue";
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                            bookMarkExists = "Acrtrue";
                    }
                }
            }
            return bookMarkExists;
        }

    }
}