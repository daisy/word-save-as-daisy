using Daisy.SaveAsDAISY.Conversion;
using System.Linq;
using System.Windows.Controls;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour DocumentMetadata.xaml
    /// </summary>
    public partial class DocumentMetadata : UserControl
    {

        public DocumentParameters Document = new DocumentParameters("");

        public DocumentMetadata()
        {
            InitializeComponent();
        }

        public DocumentMetadata(DocumentParameters d) : this()
        {
        }

        public string DocumentTitle { get; set; }
        public string DocumentAuthor { get; set; }
        public string DocumentPublisher { get; set; }
        public string DocumentIdentifier { get; set; }
        public string DocumentIdentifierScheme { get; set; }
        public string SelectedDocumentLanguage { get; set; }
        public string DocumentContributor { get; set; }
        public string DocumentSubtitle { get; set; }
        public string DocumentRights { get; set; }
        public string DocumentSummary { get; set; }
        public string DocumentSubject { get; set; }
        public string DocumentSource { get; set; }
        public string DocumentAccessibilitySummary { get; set; }

        public void LoadDocument(DocumentParameters d)
        {
            Document = d;
            DocumentTitle = Document.Title;
            DocumentAuthor = Document.Author;
            DocumentPublisher = Document.Publisher;
            DocumentIdentifier = Document.Identifier;
            DocumentIdentifierScheme = Document.IdentifierScheme;
            SelectedDocumentLanguage = Document.Languages[0];
            DocumentContributor = Document.Contributor;
            DocumentSubtitle = Document.Subtitle;
            DocumentRights = Document.Rights;
            DocumentSummary = Document.Summary;
            DocumentSubject = Document.Subject;
            DocumentSource = Document.SourceOfPagination;
            DocumentAccessibilitySummary = Document.AccessibilitySummary;
        }

        public DocumentParameters GetUpdatedDocument()
        {
            Document.Title = DocumentTitle;
            Document.Author = DocumentAuthor;
            Document.Publisher = DocumentPublisher;
            Document.Identifier = DocumentIdentifier;
            Document.IdentifierScheme = DocumentIdentifierScheme;
            Document.Languages = new System.Collections.Generic.List<string>()
            {
                SelectedDocumentLanguage,
            }.Concat(Document.Languages.Where(l => l != SelectedDocumentLanguage)).ToList();
            Document.Contributor = DocumentContributor;
            Document.Subtitle = DocumentSubtitle;
            Document.Rights = DocumentRights;
            Document.Summary = DocumentSummary;
            Document.Subject = DocumentSubject;
            Document.SourceOfPagination = DocumentSource;
            Document.AccessibilitySummary = DocumentAccessibilitySummary;
            return Document;
        }
    }
}
