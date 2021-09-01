using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Collections;
using Microsoft.Office.Core;
using MSword = Microsoft.Office.Interop.Word;

namespace Daisy.SaveAsDAISY.Addins.Word2007 {
    public partial class Abbreviation : Form {
        MSword.Document pdoc;
        ArrayList lostElements = new ArrayList();
        bool seperate;
        Int16 exists = 0;
        XmlDocument dc = new XmlDocument();

        /*Constructor which loads differently for Abbreviations/Acronyms*/
        public Abbreviation(MSword.Document doc, bool value) {
            InitializeComponent();
            pdoc = doc;
            seperate = value;

            CustomXMLParts xmlparts = pdoc.CustomXMLParts;

            if (seperate) {
                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[p].XML;
                        dc.LoadXml(partxml);

                        XmlNodeList list = dc.FirstChild.FirstChild.ChildNodes;

                        for (int j = 0; j < list.Count; j++) {
                            if (!pdoc.Bookmarks.Exists(list.Item(j).Attributes.Item(0).Value)) {
                                dc.FirstChild.FirstChild.RemoveChild(dc.FirstChild.FirstChild.ChildNodes[j]);
                                j--;
                            }
                        }

                        xmlparts[p].Delete();
                        CustomXMLPart part = pdoc.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);

                    }
                }

                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        exists = 1;
                    }
                }

                if (exists == 0) {
                    foreach (object item in pdoc.Bookmarks) {
                        if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                            object val = ((MSword.Bookmark)item).Name;
                            pdoc.Bookmarks.get_Item(ref val).Delete();
                        }
                    }
                } else {
                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            foreach (object item in pdoc.Bookmarks) {
                                if (((MSword.Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                                    String name = ((MSword.Bookmark)item).Name;
                                    NameTable nt = new NameTable();
                                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                                    nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                                    XmlNodeList node = dc.SelectNodes("//a:Item[@AbbreviationName='" + name + "']", nsManager);

                                    if (node.Count == 0) {
                                        object val = name;
                                        pdoc.Bookmarks.get_Item(ref val).Delete();
                                    }

                                    if (node.Count != 0) {
                                        if (node.Item(0).Attributes[2].Value != ((MSword.Bookmark)item).Range.Text.Trim()) {
                                            object val = name;
                                            pdoc.Bookmarks.get_Item(ref val).Delete();
                                        }

                                    }
                                }
                            }
                        }
                    }
                }


                this.Text = "Manage Abbreviation";
                this.label1.Text = "&List of Abbreviations:";


                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[p].XML;
                        dc.LoadXml(partxml);

                        XmlNodeList listAbr = dc.FirstChild.FirstChild.ChildNodes;
                        int i = 0;
                        for (int j = 0; j < listAbr.Count; j++) {
                            if (pdoc.Bookmarks.Exists(listAbr.Item(j).Attributes.Item(0).Value)) {
                                if (!lostElements.Contains(listAbr.Item(j).Attributes.Item(2).Value + "-" + listAbr.Item(j).Attributes.Item(1).Value)) {
                                    lostElements.Add(listAbr.Item(j).Attributes.Item(2).Value + "-" + listAbr.Item(j).Attributes.Item(1).Value);
                                    lBx_Abbreviation.Items.Insert(i, listAbr.Item(j).Attributes.Item(2).Value + "-" + listAbr.Item(j).Attributes.Item(1).Value);
                                    i++;
                                }
                            }
                        }
                    }
                }

                if (lBx_Abbreviation.Items.Count < 1) {
                    btn_Unmark.Enabled = false;
                }
                if (lBx_Abbreviation.Items.Count > 0) {
                    lBx_Abbreviation.SetSelected(0, true);
                }
            } else {
                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[p].XML;
                        dc.LoadXml(partxml);

                        XmlNodeList list = dc.FirstChild.FirstChild.NextSibling.ChildNodes;

                        for (int j = 0; j < list.Count; j++) {
                            if (!pdoc.Bookmarks.Exists(list.Item(j).Attributes.Item(0).Value)) {
                                dc.FirstChild.FirstChild.NextSibling.RemoveChild(dc.FirstChild.FirstChild.NextSibling.ChildNodes[j]);
                                j--;
                            }
                        }

                        xmlparts[p].Delete();
                        CustomXMLPart part = pdoc.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);
                    }
                }


                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        exists = 1;
                    }
                }

                if (exists == 0) {
                    foreach (object item in pdoc.Bookmarks) {
                        if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                            object val = ((MSword.Bookmark)item).Name;
                            pdoc.Bookmarks.get_Item(ref val).Delete();
                        }
                    }
                } else {
                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            foreach (object item in pdoc.Bookmarks) {
                                if (((MSword.Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                                    String name = ((MSword.Bookmark)item).Name;
                                    NameTable nt = new NameTable();
                                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                                    nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                                    XmlNodeList node = dc.SelectNodes("//a:Item[@AcronymName='" + name + "']", nsManager);

                                    if (node.Count == 0) {
                                        object val = name;
                                        pdoc.Bookmarks.get_Item(ref val).Delete();
                                    }

                                    if (node.Count != 0) {
                                        if (node.Item(0).Attributes[2].Value != ((MSword.Bookmark)item).Range.Text.Trim()) {
                                            object val = name;
                                            pdoc.Bookmarks.get_Item(ref val).Delete();
                                        }

                                    }
                                }
                            }
                        }
                    }
                }




                this.Text = "Manage Acronym";
                this.label1.Text = "&List of Acronyms:";

                for (int p = 1; p <= xmlparts.Count; p++) {
                    if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        String partxml = xmlparts[p].XML;
                        dc.LoadXml(partxml);

                        XmlNodeList listAcr = dc.FirstChild.FirstChild.NextSibling.ChildNodes;
                        int i = 0;
                        for (int j = 0; j < listAcr.Count; j++) {
                            if (pdoc.Bookmarks.Exists(listAcr.Item(j).Attributes.Item(0).Value)) {
                                if (!lostElements.Contains(listAcr.Item(j).Attributes.Item(2).Value + "-" + listAcr.Item(j).Attributes.Item(1).Value)) {
                                    lostElements.Add(listAcr.Item(j).Attributes.Item(2).Value + "-" + listAcr.Item(j).Attributes.Item(1).Value);
                                    lBx_Abbreviation.Items.Insert(i, listAcr.Item(j).Attributes.Item(2).Value + "-" + listAcr.Item(j).Attributes.Item(1).Value);
                                    i++;
                                }
                            }
                        }
                    }
                }

                if (lBx_Abbreviation.Items.Count < 1) {
                    btn_Unmark.Enabled = false;
                }
                if (lBx_Abbreviation.Items.Count > 0) {
                    lBx_Abbreviation.SetSelected(0, true);
                }
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e) {
            //foreach (object item in pdoc.Bookmarks)
            //{
            //    ((MSword.Bookmark)item).Range.HighlightColorIndex = MSword.WdColorIndex.wdNoHighlight;
            //}

            // pdoc.Save();
            this.Close();
        }

        /*Function which Unmark all abbreviations/Acronyms*/
        private void btn_Unmark_Click(object sender, EventArgs e) {
            CustomXMLParts xmlparts = pdoc.CustomXMLParts;

            if (seperate) {
                DialogResult dr = MessageBox.Show("Do you want to unmark this abbreviation?", "SaveAsDAISY - Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dr == DialogResult.Yes) {
                    String input = lBx_Abbreviation.SelectedItem.ToString();
                    String[] strKey = input.Split('-');


                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            XmlNodeList listAbr = dc.FirstChild.FirstChild.ChildNodes;
                            for (int j = 0; j < listAbr.Count; j++) {
                                if (listAbr.Item(j).Attributes.Item(2).Value == strKey[0] && listAbr.Item(j).Attributes.Item(1).Value == strKey[1]) {
                                    object index = listAbr.Item(j).Attributes.Item(0).Value;
                                    pdoc.Bookmarks.get_Item(ref index).Delete();
                                }
                            }
                        }
                    }

                    lostElements.Remove(input);
                    lBx_Abbreviation.Items.Remove(input);

                    if (lBx_Abbreviation.Items.Count > 0)
                        lBx_Abbreviation.SetSelected(0, true);

                    lBx_Abbreviation.Refresh();
                }
            } else {
                DialogResult dr = MessageBox.Show("Do you want to unmark this Acronym?", "SaveAsDAISY - Information", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (dr == DialogResult.Yes) {
                    String input = lBx_Abbreviation.SelectedItem.ToString();
                    String[] strKey = input.Split('-');

                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            XmlNodeList listAcr = dc.FirstChild.FirstChild.NextSibling.ChildNodes;
                            for (int j = 0; j < listAcr.Count; j++) {
                                if (listAcr.Item(j).Attributes.Item(2).Value == strKey[0] && listAcr.Item(j).Attributes.Item(1).Value == strKey[1]) {
                                    object index = listAcr.Item(j).Attributes.Item(0).Value;
                                    pdoc.Bookmarks.get_Item(ref index).Delete();
                                }
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

        private void btn_Goto_Click(object sender, EventArgs e) {
            CustomXMLParts xmlparts = pdoc.CustomXMLParts;

            if (lBx_Abbreviation.SelectedValue.ToString() != "") {
                String input = lBx_Abbreviation.SelectedValue.ToString();
                String[] strKey = input.Split('-');

                if (seperate) {
                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            XmlNodeList listAbr = dc.FirstChild.FirstChild.ChildNodes;

                            for (int j = 0; j < listAbr.Count; j++) {
                                if (listAbr.Item(j).Attributes.Item(2).Value == strKey[0] && listAbr.Item(j).Attributes.Item(1).Value == strKey[1]) {
                                    object index = listAbr.Item(j).Attributes.Item(0).Value;
                                    pdoc.Bookmarks.get_Item(ref index).Range.HighlightColorIndex = MSword.WdColorIndex.wdTurquoise;
                                }
                            }
                        }
                    }
                } else {
                    for (int p = 1; p <= xmlparts.Count; p++) {
                        if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            String partxml = xmlparts[p].XML;
                            dc.LoadXml(partxml);

                            XmlNodeList listAcr = dc.FirstChild.FirstChild.NextSibling.ChildNodes;

                            for (int j = 0; j < listAcr.Count; j++) {
                                if (listAcr.Item(j).Attributes.Item(2).Value == strKey[0] && listAcr.Item(j).Attributes.Item(1).Value == strKey[1]) {
                                    object index = listAcr.Item(j).Attributes.Item(0).Value;
                                    pdoc.Bookmarks.get_Item(ref index).Range.HighlightColorIndex = MSword.WdColorIndex.wdTurquoise;
                                }
                            }
                        }
                    }
                }

            }
        }

        private void lBx_Abbreviation_SelectedValueChanged(object sender, EventArgs e) {
            if (lBx_Abbreviation.Items.Count == 0)
                btn_Unmark.Enabled = false;
        }

        private void Abbreviation_FormClosed(object sender, FormClosedEventArgs e) {
            // pdoc.Save();
        }
    }
}