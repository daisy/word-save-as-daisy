using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Collections;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.ComponentModel;
using System.Collections.Generic;
using MSword = Microsoft.Office.Interop.Word;

namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    public partial class Abbreviation2003 : Form
    {
        MSword.Document currentDoc;
        ArrayList lostElements = new ArrayList();
        bool checkAbbrAcr;

        /// <summary>
        /// Constructor which loads differently for Abbreviations/Acronyms
        /// </summary>
        /// <param name="doc">Input Document</param>
        /// <param name="value">Checking whether call is from Abbreviation or Acronym</param>
        public Abbreviation2003(MSword.Document doc, bool value)
        {
            InitializeComponent();
            currentDoc = doc;
            checkAbbrAcr = value;

            //Getting the Custom properties of the document
            DocumentProperties xmlparts = (DocumentProperties)currentDoc.CustomDocumentProperties;

            //Deleting all custom properties whose names are starting with Abbreviation or Acronym
            //and not having bookmarks for the respective Property
            foreach (DocumentProperty docProp in xmlparts)
            {
                if (docProp.Name.StartsWith("Abbreviations") || docProp.Name.StartsWith("Acronyms"))
                {
                    if (!currentDoc.Bookmarks.Exists(docProp.Name))
                    {
                        xmlparts[docProp.Name].Delete();
                    }
                }
            }

            //checking all the bookmarks
            foreach (object item in currentDoc.Bookmarks)
            {
                bool exists = false;
                String name = ((MSword.Bookmark)item).Name;
                foreach (DocumentProperty docProp in xmlparts)
                {
                    //if Bookmark name is same as Cutom property name
                    if (docProp.Name == name)
                    {
                        exists = true;
                        String valuebook = docProp.Value.ToString();
                        String val = valuebook.Substring(0, valuebook.IndexOf("$#$"));
                        if (val != ((MSword.Bookmark)item).Range.Text.Trim())
                        {
                            ((MSword.Bookmark)item).Delete();
                        }
                    }

                }
                if (!exists)
                {
                    ((MSword.Bookmark)item).Delete();
                }
            }

            //If call is from Abbreviation button
            if (checkAbbrAcr)
            {
                //checking all Custom Properties
                foreach (DocumentProperty prop in xmlparts)
                {
                    if (currentDoc.Bookmarks.Exists(prop.Name) && prop.Name.StartsWith("Abbreviations", StringComparison.CurrentCulture))
                    {
                        int i = 0;
                        String abbValue = (String)prop.Value;
                        abbValue = abbValue.Replace("$#$", " - ");

                        if (!lostElements.Contains(abbValue))
                        {
                            lostElements.Add(abbValue);
                            lBx_Abbreviation.Items.Insert(i, abbValue);
                            i++;
                        }
                    }
                }

                this.Text = "Manage Abbreviation";
                this.label1.Text = "&List of Abbreviations:";

                if (lBx_Abbreviation.Items.Count < 1)
                {
                    btn_Unmark.Enabled = false;
                }
                if (lBx_Abbreviation.Items.Count > 0)
                {
                    lBx_Abbreviation.SetSelected(0, true);
                }
            }
            //If call is from Acronym button
            else
            {
                foreach (DocumentProperty prop in xmlparts)
                {
                    if (currentDoc.Bookmarks.Exists(prop.Name) && prop.Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                    {
                        int i = 0;
                        String acValue = (String)prop.Value;
                        acValue = acValue.Replace("$#$", " - ");

                        if (!lostElements.Contains(acValue))
                        {
                            lostElements.Add(acValue);
                            lBx_Abbreviation.Items.Insert(i, acValue);
                            i++;
                        }
                    }
                }

                this.Text = "Manage Acronym";
                this.label1.Text = "&List of Acronyms:";


                if (lBx_Abbreviation.Items.Count < 1)
                {
                    btn_Unmark.Enabled = false;
                }
                if (lBx_Abbreviation.Items.Count > 0)
                {
                    lBx_Abbreviation.SetSelected(0, true);
                }
            }
        }

        /// <summary>
        /// Function to check any Item is there or not in the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lBx_Abbreviation_SelectedValueChanged(object sender, EventArgs e)
        {
            if (lBx_Abbreviation.Items.Count == 0)
                btn_Unmark.Enabled = false;
        }


        /// <summary>
        /// Function which Unmark all abbreviations/Acronyms
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Unmark_Click(object sender, EventArgs e)
        {
            DocumentProperties xmlparts = (DocumentProperties)currentDoc.CustomDocumentProperties;

            //If call is from Abbreviation button
            if (checkAbbrAcr)
            {
                DialogResult dr = MessageBox.Show("Do you want to unmark this abbreviation?", "SaveAsDAISY - Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dr == DialogResult.Yes)
                {
                    String input = lBx_Abbreviation.SelectedItem.ToString();
                    String[] strKey = input.Split('-');

                    //Checking all the Bookmarks
                    foreach (object item in currentDoc.Bookmarks)
                    {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture) && ((MSword.Bookmark)item).Range.Text.Trim() == strKey[0].Trim())
                        {
                            object index = ((MSword.Bookmark)item).Name;
                            String name = (String)index;
                            String temp = (String)xmlparts[name].Value;
                            if (temp == input.Replace(" - ", "$#$"))
                            {
                                //Deleting the bookmark and Cutom property
                                currentDoc.Bookmarks.get_Item(ref index).Delete();
                                xmlparts[name].Delete();
                            }
                        }
                    }
                    lostElements.Remove(input);
                    lBx_Abbreviation.Items.Remove(input);
                    if (lBx_Abbreviation.Items.Count > 0)
                        lBx_Abbreviation.SetSelected(0, true);
                    lBx_Abbreviation.Refresh();
                }
            }
            //If call is from Acronym button
            else
            {
                DialogResult dr = MessageBox.Show("Do you want to unmark this Acronym?", "SaveAsDAISY - Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dr == DialogResult.Yes)
                {
                    String input = lBx_Abbreviation.SelectedItem.ToString();
                    String[] strKey = input.Split('-');

                    //Checking all the Bookmarks
                    foreach (object item in currentDoc.Bookmarks)
                    {
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture) && ((MSword.Bookmark)item).Range.Text.Trim() == strKey[0].Trim())
                        {
                            object index = ((MSword.Bookmark)item).Name;
                            String name = (String)index;
                            String temp = (String)xmlparts[name].Value;
                            if (temp == input.Replace(" - ", "$#$"))
                            {
                                //Deleting the bookmark and Cutom property
                                currentDoc.Bookmarks.get_Item(ref index).Delete();
                                xmlparts[name].Delete();
                            }
                        }
                    }
                    lostElements.Remove(input);
                    lBx_Abbreviation.Items.Remove(input);
                    if (lBx_Abbreviation.Items.Count > 0)
                        lBx_Abbreviation.SetSelected(0, true);
                    lBx_Abbreviation.Refresh();
                }
            }
        }

        /// <summary>
        /// Function to go to the particular word 
        /// based on the selected bookmark
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Goto_Click(object sender, EventArgs e)
        {
            DocumentProperties xmlparts = (DocumentProperties)currentDoc.CustomDocumentProperties;

            if (checkAbbrAcr)
            {
                String input = lBx_Abbreviation.SelectedItem.ToString();
                String[] strKey = input.Split('-');

                foreach (object item in currentDoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture) && ((MSword.Bookmark)item).Range.Text.Trim() == strKey[0].Trim())
                    {
                        object index = ((MSword.Bookmark)item).Name;
                        String name = (String)index;
                        String temp = (String)xmlparts[name].Value;
                        if (temp == input.Replace(" - ", "$#$"))
                        {
                            currentDoc.Bookmarks.get_Item(ref index).Range.HighlightColorIndex = MSword.WdColorIndex.wdTurquoise;
                        }
                    }
                }
            }
            else
            {
                String input = lBx_Abbreviation.SelectedItem.ToString();
                String[] strKey = input.Split('-');

                foreach (object item in currentDoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture) && ((MSword.Bookmark)item).Range.Text.Trim() == strKey[0].Trim())
                    {
                        object index = ((MSword.Bookmark)item).Name;
                        String name = (String)index;
                        String temp = (String)xmlparts[name].Value;
                        if (temp == input.Replace(" - ", "$#$"))
                        {
                            currentDoc.Bookmarks.get_Item(ref index).Range.HighlightColorIndex = MSword.WdColorIndex.wdTurquoise;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Function to close the Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}