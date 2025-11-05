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

        public DocumentProperties UpdatedDocumentData;
        private DocumentProperties backup;
        public bool MetadataUpdated = false;

        public Metadata(DocumentProperties data, bool isModifiable = true) : this()
        {
            UpdatedDocumentData = data;
            MetadataForm.Document = UpdatedDocumentData;
            backup = (DocumentProperties)data.Clone();
            MetadataForm.MetadataChanged += MetadataForm_MetadataChanged;
            MetadataForm.IsReadOnly = !isModifiable;
        }


        private void MetadataForm_MetadataChanged(object sender, string updatedFieldName)
        {
            SaveMetadata.IsEnabled = true; // Enable the save button when metadata is changed
        }

        private void UpdateMetadata_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MetadataUpdated = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            MetadataUpdated = false;
            UpdatedDocumentData = backup;
            Close();
        }
    }
}
