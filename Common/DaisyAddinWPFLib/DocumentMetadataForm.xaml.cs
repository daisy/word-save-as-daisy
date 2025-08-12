using Daisy.SaveAsDAISY.Conversion;
using Microsoft.Office.Interop.Word;
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
            SaveMetadata.IsEnabled = true; // No document processor, so no save operation
        }

        private void SaveMetadata_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _documentProcessor.updateDocumentMetadata(ref _document, DocumentData);
        }
    }
}
