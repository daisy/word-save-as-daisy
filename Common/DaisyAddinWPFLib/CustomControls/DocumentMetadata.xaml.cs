using Daisy.SaveAsDAISY.Conversion;
using System;
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

        private bool _isReadOnly = false;
        public bool IsReadOnly
        {
            get { return _isReadOnly; }
            set
            {
                _isReadOnly = value;
                if (_isReadOnly) {
                    TitleTextBox.IsEnabled = false;
                    AuthorTextBox.IsEnabled = false;
                    PublisherTextBox.IsEnabled = false;
                    IdentifierTextBox.IsEnabled = false;
                    IdentifierSchemeComboBox.IsEnabled = false;
                    DocumentDatePicker.IsEnabled = false;
                    ContributorTextBox.IsEnabled = false;
                    SubtitleTextBox.IsEnabled = false;
                    RightsTextBox.IsEnabled = false;
                    BookSummaryTextBox.IsEnabled = false;
                    SubjectTextBox.IsEnabled = false;
                    SourceTextBox.IsEnabled = false;
                    SourceDatePicker.IsEnabled = false;
                    AccessibilitySummaryTextBox.IsEnabled = false;
                } else {
                    TitleTextBox.IsEnabled = true;
                    AuthorTextBox.IsEnabled = true;
                    PublisherTextBox.IsEnabled = true;
                    IdentifierTextBox.IsEnabled = true;
                    IdentifierSchemeComboBox.IsEnabled = true;
                    DocumentDatePicker.IsEnabled = true;
                    ContributorTextBox.IsEnabled = true;
                    SubtitleTextBox.IsEnabled = true;
                    RightsTextBox.IsEnabled = true;
                    BookSummaryTextBox.IsEnabled = true;
                    SubjectTextBox.IsEnabled = true;
                    SourceTextBox.IsEnabled = true;
                    SourceDatePicker.IsEnabled = true;
                    AccessibilitySummaryTextBox.IsEnabled = true;
                }
            }
        }

        public delegate void MetadataChangedEventHandler(object sender, string updatedFieldName);

        public event MetadataChangedEventHandler MetadataChanged = null;

        public string DocumentTitle { get => Document.Title; set { MetadataChanged?.Invoke(this, "Title"); Document.Title = value; } }
        public string DocumentDate { get => Document.Date; set { MetadataChanged?.Invoke(this, "Date"); Document.Date = value; } }
        public string DocumentAuthor { get => Document.Author; set { MetadataChanged?.Invoke(this, "Author"); Document.Author = value; } }
        public string DocumentPublisher { get => Document.Publisher; set { MetadataChanged?.Invoke(this, "Publisher"); Document.Publisher = value; } }
        public string DocumentIdentifier { get => Document.Identifier; set { MetadataChanged?.Invoke(this, "Identifier"); Document.Identifier = value; } }
        public string DocumentIdentifierScheme { get => Document.IdentifierScheme; set { MetadataChanged?.Invoke(this, "IdentifierScheme"); Document.IdentifierScheme = value; } }
        public string DocumentContributor { get => Document.Contributor; set { MetadataChanged?.Invoke(this, "Contributor"); Document.Contributor = value; } }
        public string DocumentSubtitle { get => Document.Subtitle; set { MetadataChanged?.Invoke(this, "Subtitle"); Document.Subtitle = value; } }
        public string DocumentRights { get => Document.Rights; set { MetadataChanged?.Invoke(this, "Rights"); Document.Rights = value; } }
        public string DocumentSummary { get => Document.Summary; set { MetadataChanged?.Invoke(this, "Summary"); Document.Summary = value; } }
        public string DocumentSubject { get => Document.Subject; set { MetadataChanged?.Invoke(this, "Subject"); Document.Subject = value; } }
        public string DocumentSource { get => Document.SourceOfPagination; set { MetadataChanged?.Invoke(this, "SourceOfPagination"); Document.SourceOfPagination = value; } }
        public string DocumentAccessibilitySummary { get => Document.AccessibilitySummary; set { MetadataChanged?.Invoke(this, "AccessibilitySummary"); Document.AccessibilitySummary = value; } }
        public string DocumentSourceDate { get => Document.SourceDate; set { MetadataChanged?.Invoke(this, "SourceDate"); Document.SourceDate = value; } }



    }
}
