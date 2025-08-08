using Daisy.SaveAsDAISY.Conversion;
using System.Windows.Forms;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using UserControl = System.Windows.Controls.UserControl;
using System.IO;
using System;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour PathParameter.xaml
    /// </summary>
    public partial class PathParameter : UserControl
    {
        #region fields
        
        public ScriptParameter BoundParameter = null;
        private string _parameterName = string.Empty;
        private string _parameterValue = string.Empty;
        private string _parameterDescription = string.Empty;


        private bool _isFile = false; // Default to directory selection
        public bool IsFile { get => BoundParameter != null ? (BoundParameter.ParameterData as PathData).IsFile : _isFile; set => _isFile = value; } // Default to directory selection

        private bool _isInput = false; // Default to output selection
        public bool IsInput { get => BoundParameter != null ? (BoundParameter.ParameterData as PathData).IsInput : _isInput; set => _isInput = value; } // default to output selection

        private string _fileExtension = string.Empty; // Default to no file extension
        public string FileExtension { get => BoundParameter != null ? (BoundParameter.ParameterData as PathData).FileExtension : _fileExtension; set => _fileExtension = value; } // default to output selection
        public string InitialDirectory { get; set; } = string.Empty;
        public string ParameterName {
            get => BoundParameter != null ? BoundParameter.NiceName : _parameterName;
            set => _parameterName = value;
        }
        public string ParameterValue {
            get => BoundParameter != null ? (string)BoundParameter.Value : _parameterValue;
            set
            {
                _parameterValue = value;
                if (BoundParameter != null) {
                    BoundParameter.Value = value;
                }
            }
        }

        public string ParameterDescription {
            get => BoundParameter != null ? BoundParameter.Description : _parameterDescription;
            set =>  _parameterDescription = value;
        }

        #endregion

        public PathParameter()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public PathParameter(ScriptParameter p)
           : this()
        {
            BoundParameter = p;
            PathData linkedPathData = p.ParameterData as PathData;
            IsFile = linkedPathData.IsFile;
            IsInput = linkedPathData.IsInput;
            FileExtension = linkedPathData.FileExtension;
            try {
                InitialDirectory = linkedPathData.Value != null ? linkedPathData.Value.ToString() : "";
            }
            catch (Exception ex) {
                InitialDirectory = "";
                System.Diagnostics.Debug.WriteLine("Error retrieving initial directory: " + ex.Message);
            }
        }


        private void BrowseButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (IsFile) {
                if (IsInput) {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    if (!string.IsNullOrEmpty(FileExtension)) {
                        openFileDialog.Filter = FileExtension + "|*" + FileExtension;
                    }

                    if (InitialDirectory != "" && File.Exists(InitialDirectory)) {
                        openFileDialog.InitialDirectory = Path.GetDirectoryName(InitialDirectory);
                    }
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == true) {
                        ParameterValue = openFileDialog.FileName;
                    }
                } else {
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    if (InitialDirectory != "" && File.Exists(InitialDirectory)) {
                        saveFileDialog.InitialDirectory = Path.GetDirectoryName(InitialDirectory);
                    }
                    if (saveFileDialog.ShowDialog() == true) {
                        ParameterValue = saveFileDialog.FileName;
                    }
                }
            } else {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                if (InitialDirectory != "" && File.Exists(InitialDirectory)) {
                    folderBrowserDialog.SelectedPath = Path.GetDirectoryName(InitialDirectory);
                }
                if(folderBrowserDialog.ShowDialog() == DialogResult.OK) {
                    ParameterValue = folderBrowserDialog.SelectedPath;
                }
            }
            PathTextBox.Text = ParameterValue;
        }

    }
}
