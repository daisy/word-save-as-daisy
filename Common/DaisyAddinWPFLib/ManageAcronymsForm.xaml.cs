using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;

namespace Daisy.SaveAsDAISY.WPF
{
    using MSWord = Microsoft.Office.Interop.Word;

    /// <summary>
    /// Logique d'interaction pour ManageAcronymsForm.xaml
    /// </summary>
    public partial class ManageAcronymsForm : System.Windows.Window
    {
        private class AcronymItem
        {
            public string AcronymName { get; set; }
            public string OriginalText { get; set; }
            public string FullAcr { get; set; }
            public bool PronouncedAsWord { get; set; }
        }


        private List<AcronymItem> Acronyms { get; set; } = new List<AcronymItem>();

        public ManageAcronymsForm()
        {
            InitializeComponent();
        }

        MSWord.Document CurrentDocument { get; set; }
        CustomXMLPart xmlPart = null;
        XmlDocument dc = new XmlDocument();

        public ManageAcronymsForm(Document current) : this()
        {
            CurrentDocument = current;
            // Revamped from previous version :
            // Cleaning the abbreviation from the document if they have been removed
            CustomXMLParts xmlparts = CurrentDocument.CustomXMLParts;
            for (int p = 1; p <= xmlparts.Count && xmlPart == null; p++) {
                if (xmlparts[p].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                    xmlPart = xmlparts[p];
                    string partxml = xmlPart.XML;
                    dc.LoadXml(partxml);
                    XmlElement elmManage = (XmlElement)dc.FirstChild;
                    XmlElement elmAcronyms = null;
                    bool hasChanged = false;
                    for (int j = 0; j < elmManage.ChildNodes.Count; j++) {
                        if (elmManage.ChildNodes[j].Name == "Acronyms") {
                            elmAcronyms = (XmlElement)elmManage.ChildNodes[j];
                        }
                    }
                    if (elmAcronyms != null) {
                        XmlNodeList list = elmAcronyms.ChildNodes;
                        for (int j = list.Count - 1; j >= 0; j--) {
                            AcronymItem item = new AcronymItem()
                            {
                                AcronymName = list.Item(j).Attributes.GetNamedItem("AcronymName")?.Value,
                                OriginalText = list.Item(j).Attributes.GetNamedItem("OriginalText")?.Value,
                                FullAcr = list.Item(j).Attributes.GetNamedItem("FullAcr")?.Value,
                                PronouncedAsWord = list.Item(j).Attributes.GetNamedItem("AcronymName")?.Value.StartsWith("AcronymsYes") == true
                            };
                            if (!CurrentDocument.Bookmarks.Exists(item.AcronymName)) {
                                // Abbreviation was removed from the document bookmarks, remove it from the XML part
                                elmAcronyms.RemoveChild(elmAcronyms.ChildNodes[j]);
                                hasChanged = true;
                            } else {
                                Acronyms.Add(item);
                            }
                        }
                        foreach (object item in CurrentDocument.Bookmarks) {
                            if (((Bookmark)item).Name.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                                string name = ((Bookmark)item).Name;
                                AcronymItem found = Acronyms.First((x) => x.AcronymName == name);
                                if (found == null || (found != null && found.OriginalText != ((Bookmark)item).Range.Text.Trim())) {
                                    // The abbreviation is not in the XML part, remove the bookmark
                                    object val = name;
                                    CurrentDocument.Bookmarks.get_Item(ref val).Delete();
                                    Acronyms.Remove(found);
                                }
                            }
                        }
                        if (Acronyms.Count == 0) {
                            // If no acronyms are left, remove the Acronyms element
                            elmManage.RemoveChild(elmAcronyms);
                            hasChanged = true;
                        }
                    }
                    if (hasChanged) {
                        // If the XML part has been modified, replace it in the document
                        xmlPart.Delete();
                        xmlPart = CurrentDocument.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);
                    }
                }
            }
            Acronyms.Reverse();
            AcronymsList.ItemsSource = Acronyms;
        }

        private void Acronyms_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectInDocument.IsEnabled = AcronymsList.SelectedIndex >= 0;
            UnmarkAcronym.IsEnabled = AcronymsList.SelectedIndex >= 0;
            //if (AbbreviationsList.SelectedIndex >= 0) {
            //    AbbreviationItem item = (AbbreviationItem)AbbreviationsList.SelectedItem;
            //    object index = item.AbbreviationName;
            //    Range rng = CurrentDocument.Bookmarks.get_Item(ref index).Range;
            //    rng.Select();
            //    rng.Select();
            //}

        }

        private void CloseForm_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SelectInDocument_Click(object sender, RoutedEventArgs e)
        {
            CustomXMLParts xmlparts = CurrentDocument.CustomXMLParts;
            AcronymItem item = (AcronymItem)AcronymsList.SelectedItem;
            object index = item.AcronymName;
            Range rng = CurrentDocument.Bookmarks.get_Item(ref index).Range;
            rng.Select();
            //rng.HighlightColorIndex = MSWord.WdColorIndex.wdTurquoise;
            rng.Select();
        }

        private void UnmarkAcronym_Click(object sender, RoutedEventArgs e)
        {
            // Remove abreviation from list and update xml part
            AcronymItem item = (AcronymItem)AcronymsList.SelectedItem;
            // Remove bookmark from the document
            if (CurrentDocument.Bookmarks.Exists(item.AcronymName)) {
                object val = item.AcronymName;
                // Remove the bookmark from the document
                CurrentDocument.Bookmarks.get_Item(ref val).Delete();
            }

            // Remove abbreviation from the XML part
            if (xmlPart != null) {
                XmlElement elmManage = (XmlElement)dc.FirstChild;
                XmlElement elmAcronyms = null;
                for (int j = 0; j < elmManage.ChildNodes.Count; j++) {
                    if (elmManage.ChildNodes[j].Name == "Acronyms") {
                        elmAcronyms = (XmlElement)elmManage.ChildNodes[j];
                    }
                }
                if (elmAcronyms == null) {
                    elmAcronyms = dc.CreateElement("Acronyms");
                    elmManage.AppendChild(elmAcronyms);
                }

                XmlNodeList list = elmAcronyms.ChildNodes;
                for (int j = list.Count - 1; j >= 0; j--) {
                    if (list.Item(j).Attributes.GetNamedItem("AcronymName")?.Value == item.AcronymName) {
                        elmAcronyms.RemoveChild(elmAcronyms.ChildNodes[j]);
                        break;
                    }
                }
                // Replace the custom XML part with the cleaned one
                xmlPart.Delete();
                xmlPart = CurrentDocument.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);
            }
            // Remove from the list
            Acronyms.Remove(item);
            AcronymsList.ItemsSource = Acronyms;
            AcronymsList.Items.Refresh();
            AcronymsList.SelectedIndex = -1;
        }
    }
}
