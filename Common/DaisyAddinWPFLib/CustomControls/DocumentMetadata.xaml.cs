using Daisy.SaveAsDAISY.Conversion;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour DocumentMetadata.xaml
    /// </summary>
    public partial class DocumentMetadata : UserControl
    {

        public DocumentProperties Document = new DocumentProperties("");


        public DocumentMetadata()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public DocumentMetadata(DocumentProperties d)
        {
            Document = d;
            InitializeComponent();
            this.DataContext = this;
        }

        //public void BindTo(DocumentProperties d)
        //{
        //    Document = d;
        //}

        public string DocumentTitle { get => Document.Title; set => Document.Title = value; }
        public string DocumentDate { get => Document.Date; set => Document.Date = value; }
        public string DocumentAuthor { get => Document.Author; set => Document.Author = value; }
        public string DocumentPublisher { get => Document.Publisher; set => Document.Publisher = value; }
        public string DocumentIdentifier { get => Document.Identifier; set => Document.Identifier = value; }
        public string DocumentIdentifierScheme { get => Document.IdentifierScheme; set => Document.IdentifierScheme = value; }
        public string DocumentContributor { get => Document.Contributor; set => Document.Contributor = value; }
        public string DocumentSubtitle { get => Document.Subtitle; set => Document.Subtitle = value; }
        public string DocumentRights { get => Document.Rights; set => Document.Rights = value; }
        public string DocumentSummary { get => Document.Summary; set => Document.Summary = value; }
        public string DocumentSubject { get => Document.Subject; set => Document.Subject = value; }
        public string DocumentSource { get => Document.SourceOfPagination; set => Document.SourceOfPagination = value; }
        public string DocumentAccessibilitySummary { get => Document.AccessibilitySummary; set => Document.AccessibilitySummary = value; }

        public string DocumentSourceDate { get => Document.SourceDate; set => Document.SourceDate = value; }



    }
}
