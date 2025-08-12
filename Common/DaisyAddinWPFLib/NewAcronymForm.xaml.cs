using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Xml;


namespace Daisy.SaveAsDAISY.WPF
{
    using MSword = Microsoft.Office.Interop.Word;

    /// <summary>
    /// Logique d'interaction pour AcronymForm.xaml
    /// </summary>
    public partial class NewAcronymForm : System.Windows.Window
    {
        public NewAcronymForm()
        {
            InitializeComponent();
            DataContext = this;
        }

        public NewAcronymForm(Document current) : this()
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
                         && (item.Name.StartsWith("Acronyms", StringComparison.CurrentCulture)
                         || item.Name.StartsWith("Acronyms", StringComparison.CurrentCulture))
                    ) {
                        return true;
                    }
                }
            }
            return false;
        }


        private void SaveAcronym_Click(object sender, RoutedEventArgs e)
        {
            try {
                if (currentDocument.ProtectionType == WdProtectionType.wdNoProtection) {
                    int customXmlPartIndex = 0;
                    CustomXMLParts xmlParts = currentDocument.CustomXMLParts;
                    XmlElement elmtAcronyms = null;
                    XmlDocument customXml = new XmlDocument();
                    //Checking whether CustomXML is there or not in the Current Document
                    for (int i = 1; i <= xmlParts.Count; i++) {
                        if (xmlParts[i].NamespaceURI == "http://Daisy-OpenXML/customxml") {
                            customXmlPartIndex = i;
                            string partxml = xmlParts[i].XML;
                            customXml.LoadXml(partxml);
                            XmlElement elmtManage = (XmlElement)customXml.FirstChild;
                            for (int j = 0; j < elmtManage.ChildNodes.Count; j++) {
                                if (elmtManage.ChildNodes[j].Name == "Acronyms") {
                                    elmtAcronyms = (XmlElement)elmtManage.ChildNodes[j];
                                    break;
                                }
                            }
                            if (elmtAcronyms == null) {
                                elmtAcronyms = customXml.CreateElement("Acronyms");
                                elmtManage.AppendChild(elmtAcronyms);
                            }
                        }
                    }
                    if (customXmlPartIndex == 0) {
                        // Prepare the custom XML document to be embeded in the current document
                        XmlElement elmtManage = customXml.CreateElement("Manage");
                        customXml.AppendChild(elmtManage);
                        elmtManage.SetAttribute("xmlns", "http://Daisy-OpenXML/customxml");
                        // Create the Acronyms storage element
                        elmtAcronyms = customXml.CreateElement("Acronyms");
                        elmtManage.AppendChild(elmtAcronyms);
                    }

                    // If only the current selection is to be marked as occurence
                    if (ApplyEverywhere.IsChecked != true) {
                        object val = System.Reflection.Missing.Value;
                        string nameBookmark = "Acronyms" 
                            + (PronouncedAsWord.IsChecked == true ? "Yes" : "No") 
                            + GenerateId().ToString();
                        //Adding Bookmark for the selected Text
                        currentDocument.Application.Selection.Bookmarks.Add(nameBookmark, ref val);

                        //Updating the CustomXML
                        XmlElement elmtItem = customXml.CreateElement("Item");
                        elmtItem.SetAttribute("AcronymName", nameBookmark);
                        elmtItem.SetAttribute("FullAcr", FullFormTextBox.Text.TrimEnd());
                        elmtItem.SetAttribute("OriginalText", currentDocument.Application.Selection.Text.Trim());
                        elmtAcronyms.AppendChild(elmtItem);

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

                        //Applying Acronym for all the occurrences of the selected Word
                        while (fndDoc.Found) {
                            if (!IsMarkedAsAbbreviationOrAcronym(rngDoc)) {
                                string nameBookmark = "Acronyms"
                                    + (PronouncedAsWord.IsChecked == true ? "Yes" : "No")
                                    + GenerateId().ToString();
                                rngDoc.Bookmarks.Add(nameBookmark, ref missing);

                                XmlElement elmtItem = customXml.CreateElement("Item");
                                elmtItem.SetAttribute("AcronymName", nameBookmark);
                                elmtItem.SetAttribute("FullAcr", FullFormTextBox.Text.TrimEnd());
                                elmtItem.SetAttribute("OriginalText", currentDocument.Application.Selection.Text.Trim());

                                elmtAcronyms.AppendChild(elmtItem);
                            }
                            fndDoc.Execute(
                                    ref missing, ref MatchCase, ref wholeword,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing);
                        }
                    }
                    if (customXmlPartIndex > 0) {
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

        private void CancelAcronym_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

    }
}
