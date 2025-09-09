using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Xml;

namespace Daisy.SaveAsDAISY.WPF
{
    using MSword = Microsoft.Office.Interop.Word;

    /// <summary>
    /// Logique d'interaction pour NewAbreviationForm.xaml
    /// </summary>
    public partial class NewAbreviationForm : System.Windows.Window
    {
        public NewAbreviationForm()
        {
            InitializeComponent();
            DataContext = this;
        }

        public NewAbreviationForm(Document current) : this()
        {
            currentDocument = current;
            Selection = currentDocument.Application.Selection.Text.Trim();
        }

        private Document currentDocument;

        public string Selection { get; private set; }

        private long GenerateId()
        {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        private bool IsMarkedAsAbbreviationOrAcronym(MSword.Range rngDoc)
        {
            if (rngDoc.Bookmarks.Count > 0) {
                foreach (MSword.Bookmark item in currentDocument.Bookmarks) {
                    if (item.Range.Text.Trim() == rngDoc.Text.Trim()
                         && (item.Name.StartsWith("Abbreviations", StringComparison.CurrentCulture)
                         || item.Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                    ) {
                        return true;
                    }
                }
            }
            return false;
        }

        private void SaveAbbreviation_Click(object sender, RoutedEventArgs e)
        {
            try {
                if (currentDocument.ProtectionType == WdProtectionType.wdNoProtection) {
                    int customXmlPartIndex = 0;
                    CustomXMLParts xmlParts = currentDocument.CustomXMLParts;
                    XmlElement elmtAbbreviations = null;
                    XmlDocument customXml = new XmlDocument();
                    //Checking whether CustomXML is there or not in the Current Document
                    for (int i = 1; i <= xmlParts.Count; i++) {
                        if (xmlParts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            customXmlPartIndex = i;
                            string partxml = xmlParts[i].XML;
                            customXml.LoadXml(partxml);
                            XmlElement elmtManage = (XmlElement)customXml.FirstChild;
                            for (int j = 0; j < elmtManage.ChildNodes.Count; j++) {
                                if (elmtManage.ChildNodes[j].Name == "Abbreviations") {
                                    elmtAbbreviations = (XmlElement)elmtManage.ChildNodes[j];
                                    break;
                                }
                            }
                            if(elmtAbbreviations == null) {
                                elmtAbbreviations = customXml.CreateElement("Abbreviations", "http://Daisy-OpenXML/customxml");
                                elmtManage.AppendChild(elmtAbbreviations);
                            }
                        }
                    }
                    if (customXmlPartIndex == 0) {
                        // Prepare the custom XML document to be embeded in the current document
                        XmlElement elmtManage = customXml.CreateElement("Manage", "http://Daisy-OpenXML/customxml");
                        customXml.AppendChild(elmtManage);
                        elmtManage.SetAttribute("xmlns", "http://Daisy-OpenXML/customxml");
                        // Create the Abbreviations storage element
                        elmtAbbreviations = customXml.CreateElement("Abbreviations", "http://Daisy-OpenXML/customxml");
                        elmtManage.AppendChild(elmtAbbreviations);
                        // Also create the acronym elements manage by the same custom XML
                        // Will be done
                        //XmlElement elmtAcronyms = customXml.CreateElement("Acronyms");
                        //elmtManage.AppendChild(elmtAcronyms);
                    }
                    
                    // If only the current selection is to be marked as occurence
                    if (ApplyEverywhere.IsChecked != true) {
                        object val = System.Reflection.Missing.Value;
                        string nameAbbr = "Abbreviations" + GenerateId().ToString();
                        //Adding Bookmark for the selected Text
                        currentDocument.Application.Selection.Bookmarks.Add(nameAbbr, ref val);

                        //Updating the CustomXML
                        XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                        elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                        elmtItem.SetAttribute("FullAbbr", FullFormTextBox.Text.TrimEnd());
                        elmtItem.SetAttribute("OriginalText", currentDocument.Application.Selection.Text.Trim());
                        elmtAbbreviations.AppendChild(elmtItem);

                    } else {
                        object missing = System.Type.Missing;
                        object startDoc = 0;
                        object endDoc = currentDocument.Characters.Count;
                        Range rngDoc = currentDocument.Range(ref startDoc, ref endDoc);
                        Find fndDoc = rngDoc.Find;

                        fndDoc.ClearFormatting();
                        fndDoc.Forward = true;
                        fndDoc.Text = currentDocument.Application.Selection.Text.Trim();
                        object wholeword = true;
                        object MatchCase = true;

                        fndDoc.Execute(ref missing, ref MatchCase, ref wholeword,
                          ref missing, ref missing, ref missing,
                          ref missing, ref missing, ref missing, ref missing, ref missing,
                          ref missing, ref missing, ref missing,
                          ref missing);

                        //Applying Abbreviation for all the occurrences of the selected Word
                        while (fndDoc.Found) {
                            if (!IsMarkedAsAbbreviationOrAcronym(rngDoc)) {
                                string nameAbbr = "Abbreviations" + GenerateId().ToString();
                                rngDoc.Bookmarks.Add(nameAbbr, ref missing);

                                XmlElement elmtItem = customXml.CreateElement("Item", "http://Daisy-OpenXML/customxml");
                                elmtItem.SetAttribute("AbbreviationName", nameAbbr);
                                elmtItem.SetAttribute("FullAbbr", FullFormTextBox.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDocument.Application.Selection.Text.Trim());

                                elmtAbbreviations.AppendChild(elmtItem);
                            }
                            fndDoc.Execute(
                                    ref missing, ref MatchCase, ref wholeword,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing);
                        }
                    }
                    if(customXmlPartIndex > 0) {
                        // If CustomXML is already present in current Document, we replace it
                        xmlParts[customXmlPartIndex].Delete();
                    }
                    //Adding the CustomXml to the Current Document
                    CustomXMLPart part = currentDocument.CustomXMLParts.Add(customXml.InnerXml, System.Reflection.Missing.Value);
                    Close();

                } else {
                    MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CancelAbbreviation_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
