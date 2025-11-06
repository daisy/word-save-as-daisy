using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
    /// Logique d'interaction pour ManageAbbreviationsForm.xaml
    /// </summary>
    public partial class ManageAbbreviationsForm : System.Windows.Window
    {
        private class AbbreviationItem
        {
            public string AbbreviationName { get; set; }
            public string OriginalText { get; set; }
            public string FullAbbr { get; set; }
        }

        private List<AbbreviationItem> Abbreviations { get; set; } = new List<AbbreviationItem>();

        public ManageAbbreviationsForm()
        {
            InitializeComponent();
            SelectInDocument.IsEnabled = AbbreviationsList.SelectedIndex >= 0;
            UnmarkAbbreviation.IsEnabled = AbbreviationsList.SelectedIndex >= 0;
        }

        MSWord.Document CurrentDocument { get; set; }
        CustomXMLPart xmlPart = null;
        XmlDocument dc = new XmlDocument();

        public ManageAbbreviationsForm(Document current) : this()
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
                    XmlElement elmAbbreviations = null;
                    bool hasChanged = false;
                    for (int j = 0; j < elmManage.ChildNodes.Count; j++) {
                        if (elmManage.ChildNodes[j].Name == "Abbreviations") {
                            elmAbbreviations = (XmlElement)elmManage.ChildNodes[j];
                        }
                    }
                    if (elmAbbreviations != null) {
                        XmlNodeList list = elmAbbreviations.ChildNodes;
                        for (int j = list.Count - 1; j >= 0; j--) {
                            AbbreviationItem item = new AbbreviationItem()
                            {
                                AbbreviationName = list.Item(j).Attributes.GetNamedItem("AbbreviationName")?.Value,
                                OriginalText = list.Item(j).Attributes.GetNamedItem("OriginalText")?.Value,
                                FullAbbr = list.Item(j).Attributes.GetNamedItem("FullAbbr")?.Value
                            };
                            if (!CurrentDocument.Bookmarks.Exists(item.AbbreviationName)) {
                                // Abbreviation was removed from the document bookmarks, remove it from the XML part
                                elmAbbreviations.RemoveChild(elmAbbreviations.ChildNodes[j]);
                                hasChanged = true;
                            } else {
                                Abbreviations.Add(item);
                            }
                        }

                        foreach (object item in CurrentDocument.Bookmarks) {
                            if (((Bookmark)item).Name.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                                string name = ((Bookmark)item).Name;
                                AbbreviationItem found = Abbreviations.First((x) => x.AbbreviationName == name);
                                if (found == null || (found != null && found.OriginalText != ((Bookmark)item).Range.Text.Trim())) {
                                    // The abbreviation is not in the XML part, remove the bookmark
                                    object val = name;
                                    CurrentDocument.Bookmarks.get_Item(ref val).Delete();
                                    Abbreviations.Remove(found);
                                }
                            }
                        }
                        if (Abbreviations.Count == 0) {
                            // If no acronyms are left, remove the Acronyms element
                            elmManage.RemoveChild(elmAbbreviations);
                            hasChanged = true;
                        }
                        if (hasChanged) {
                            // If the XML part has been modified, replace it in the document
                            xmlPart.Delete();
                            xmlPart = CurrentDocument.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);
                        }
                    }
                }
            }
            Abbreviations.Reverse();
            AbbreviationsList.ItemsSource = Abbreviations;
            if (AbbreviationsList.Items.Count > 0) {
                AbbreviationsList.SelectedIndex = 0;
            }
        }

        private void Abbreviations_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectInDocument.IsEnabled = AbbreviationsList.SelectedIndex >= 0;
            UnmarkAbbreviation.IsEnabled = AbbreviationsList.SelectedIndex >= 0;
            PreviousAbbreviation.IsEnabled = AbbreviationsList.SelectedIndex > 0;
            NextAbbreviation.IsEnabled = AbbreviationsList.SelectedIndex >= 0 && AbbreviationsList.SelectedIndex < AbbreviationsList.Items.Count - 1;
            if (AbbreviationsList.SelectedIndex >= 0) {
                AbbreviationItem item = (AbbreviationItem)AbbreviationsList.SelectedItem;
                object index = item.AbbreviationName;
                Range rng = CurrentDocument.Bookmarks.get_Item(ref index).Range;
                rng.Select();
                //rng.HighlightColorIndex = MSWord.WdColorIndex.wdTurquoise;
                rng.Select();
            }

        }

        private void CloseForm_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void SelectInDocument_Click(object sender, RoutedEventArgs e)
        {
            AbbreviationItem item = (AbbreviationItem)AbbreviationsList.SelectedItem;
            object index = item.AbbreviationName;
            Range rng = CurrentDocument.Bookmarks.get_Item(ref index).Range;
            rng.Select();
            //rng.HighlightColorIndex = MSWord.WdColorIndex.wdTurquoise;
            rng.Select();
        }

        private void UnmarkAbbreviation_Click(object sender, RoutedEventArgs e)
        {
            // 2025/10/13: remarks from tester -
            // > If an acronym is accidentally unmarked, there is currently no option to re-add it to the same list.
            // > A new list has to be created instead.
            // Not sure how to handle the "re-addition" / "undo" of an abbreviation,
            // so I use a confirm message box for now

            // Remove abreviation from list and update xml part
            AbbreviationItem item = (AbbreviationItem)AbbreviationsList.SelectedItem;
            if(MessageBox.Show(
                    "Are you sure you want to unmark the abbreviation '" + item.OriginalText + "' ?",
                    "Confirm unmarking",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question
                ) != MessageBoxResult.Yes
            ) {
                // early return if user does not confirm
                return;
            }
            // Remove bookmark from the document
            if (CurrentDocument.Bookmarks.Exists(item.AbbreviationName)) {
                object val = item.AbbreviationName;
                // Remove the bookmark from the document
                CurrentDocument.Bookmarks.get_Item(ref val).Delete();
            }

            // Remove abbreviation from the XML part
            if (xmlPart != null) {
                XmlElement elmManage = (XmlElement)dc.FirstChild;
                XmlElement elmAbbreviations = null;
                for (int j = 0; j < elmManage.ChildNodes.Count; j++) {
                    if (elmManage.ChildNodes[j].Name == "Abbreviations") {
                        elmAbbreviations = (XmlElement)elmManage.ChildNodes[j];
                    }
                }
                if (elmAbbreviations == null) {
                    elmAbbreviations = dc.CreateElement("Abbreviations");
                    elmManage.AppendChild(elmAbbreviations);
                }
                
                XmlNodeList list = elmAbbreviations.ChildNodes;
                for (int j = list.Count - 1; j >= 0; j--) {
                    if (list.Item(j).Attributes.GetNamedItem("AbbreviationName")?.Value == item.AbbreviationName) {
                        elmAbbreviations.RemoveChild(elmAbbreviations.ChildNodes[j]);
                        break;
                    }
                }
                // Replace the custom XML part with the cleaned one
                xmlPart.Delete();
                xmlPart = CurrentDocument.CustomXMLParts.Add(dc.InnerXml, System.Reflection.Missing.Value);
            }
            // Remove from the list
            Abbreviations.Remove(item);
            AbbreviationsList.ItemsSource = Abbreviations;
            AbbreviationsList.Items.Refresh();
            AbbreviationsList.SelectedIndex = -1;
        }


        private void PreviousAbbreviation_Click(object sender, RoutedEventArgs e)
        {
            AbbreviationsList.SelectedIndex--;
        }

        private void NextAbbreviation_Click(object sender, RoutedEventArgs e)
        {
            AbbreviationsList.SelectedIndex++;
        }
    }
}
