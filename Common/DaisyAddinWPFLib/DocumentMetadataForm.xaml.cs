using Daisy.SaveAsDAISY.Conversion;
using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using Window = System.Windows.Window;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour DocumentMetadata.xaml
    /// </summary>
    public partial class Metadata : Window
    {
        public Metadata()
        {
            InitializeComponent();
        }
        IDocumentPreprocessor _documentProcessor = null;
        object _document = null;

        public DocumentProperties DocumentData;

        public Metadata(IDocumentPreprocessor proc, ref object document) : this()
        {
            _documentProcessor = proc;
            _document = document;
            DocumentData = proc.loadDocumentParameters(ref _document);
            MetadataForm.Document = DocumentData;
            MetadataForm.MetadataChanged += MetadataForm_MetadataChanged;
            
        }

        private void MetadataForm_MetadataChanged(object sender, string updatedFieldName)
        {
            SaveMetadata.IsEnabled = true; // Enable the save button when metadata is changed
        }

        private void UpdateMetadata_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try {
                _documentProcessor.updateDocumentMetadata(ref _document, DocumentData);
            } catch (Exception ex)
            {
                MessageBox.Show(
                    "Error updating metadata: " + ex.Message 
                    + "\r\nPlease try saving the document on another place on your system before editing metadatas.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);

            }
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
