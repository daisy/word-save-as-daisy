using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using System;
using System.IO;
using System.Linq;
using System.Windows;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Conversion form that, given a document (and a processor class to parse it) and a set of predefined conversion parameters,
    /// will allow a user to modify a set of conversion parameters (DAISY Pipeline 2 scripts inputs and options)
    /// </summary>
    public partial class ConversionParametersForm : Window
    {
        public ConversionParameters UpdatedConversionParameters { get; private set; }

        public DocumentProperties DocumentProps { get; private set; } = null;

        //private IDocumentPreprocessor _proc = null;
        //private object _document = null;

        public ConversionParametersForm()
        {
            InitializeComponent();
            DocumentMetadata.DetailedMetadataSection.Expanded += DetailedMetadataSection_Expanded;
            DocumentMetadata.DetailedMetadataSection.Collapsed += DetailedMetadataSection_Expanded;
        }

        public ConversionParametersForm(Script _script, DocumentProperties preloadedProps = null) : this()
        {
            if(_script != null) {
                PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();
                DestinationControl.BoundParameter = _script.Parameters["output"];
                DestinationControl.ParameterValue = prepopulateDaisyOutput != null
                                      ? prepopulateDaisyOutput.OutputPath
                                      : preloadedProps != null 
                                            ? Path.GetDirectoryName(preloadedProps.InputPath)
                                            :  ConverterSettings.Instance.ResultsFolder;

                if (preloadedProps != null) {
                    DocumentProps = preloadedProps;
                    DocumentMetadata.Document = DocumentProps;
                    if (DocumentProps.Languages.Count > 0 && _script.Parameters.ContainsKey("language")) {
                        ScriptParameter sp = _script.Parameters["language"];
                        _script.Parameters["language"] = new ScriptParameter(
                            sp.Name,
                            sp.NiceName,
                            new EnumData(DocumentProps.Languages.ToDictionary(s => s, s => (object)s), DocumentProps.Languages[0]),
                            sp.IsParameterRequired,
                            sp.Description,
                            sp.IsParameterDisplayed,
                            sp.Direction
                            );
                    }
                }

                foreach (var item in _script.Parameters.Where(kv => kv.Key != "input" && kv.Key != "output")) {
                    ScriptParameter sp = item.Value;
                    ParameterDataBase parameter = sp.ParameterData;

                    if (sp.IsParameterRequired || sp.IsParameterDisplayed) {
                        if (parameter is BoolData) {
                            ScriptOptions.Children.Add(new CustomControls.BooleanParameter(sp));
                        } else if (parameter is EnumData) {
                            ScriptOptions.Children.Add(new CustomControls.EnumParameter(sp));
                        } else if (parameter is StringData) {
                            ScriptOptions.Children.Add(new CustomControls.StringParameter(sp));
                        } else if (parameter is PathData) {
                            ScriptOptions.Children.Add(new CustomControls.PathParameter(sp));
                        } else if (parameter is IntegerData) {
                            ScriptOptions.Children.Add(new CustomControls.IntegerParameter(sp));
                        } else {
                            ScriptOptions.Children.Add(new CustomControls.StringParameter(sp));
                        }
                    }
                }
                DataContext = this;
                ConversionTitle = _script.NiceName;
                this.UpdatedConversionParameters = new ConversionParameters()
                {
                    PipelineScript = _script,
                    NameValidator = StringValidator.DTBookXMLFileNameFormat,
                };
            }
        }

        

        public ConversionParametersForm(IDocumentPreprocessor proc, ref object document, ConversionParameters c, DocumentProperties preloadedProps = null) 
            : this(c.PipelineScript, preloadedProps ?? proc.loadDocumentParameters(ref document))
        {
            //_proc = proc;
            //_document = document;
            this.UpdatedConversionParameters = c;
           
        }

        public string ConversionTitle { get; set; }

        private bool _performConversion = false;
        
        public new bool ShowDialog()
        {
            base.ShowDialog();
            return _performConversion;
           
            
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (DocumentMetadata.TitleTextBox.Text.TrimEnd() == "") {
                MessageBox.Show(Labels.EnterTitle, Labels.Error, MessageBoxButton.OK, MessageBoxImage.Error);
                DocumentMetadata.TitleTextBox.Focus();
                return;
            }
            // Save the data in the document properties
            //if (_proc != null && _document != null && DocumentProps != null) {
            //    _proc.updateDocumentMetadata(ref _document, DocumentProps);

            //}
            // Update script parameters with the document properties
            foreach (var kv in DocumentProps.ParametersValues) {
                if (UpdatedConversionParameters.PipelineScript.Parameters.ContainsKey(kv.Key)) {
                    UpdatedConversionParameters.PipelineScript.Parameters[kv.Key].Value = kv.Value;
                }
            }
            // Required for now to open the result global folder but not sure it is the best place to do it
            // I also have the option to open the final result folder directly from the conversion script
            UpdatedConversionParameters.OutputPath = DestinationControl.ParameterValue;
            // TODO : keeping this for now but not sure if it would not be better to keep a "default" output path in addin settings
            new PrepopulateDaisyOutput(UpdatedConversionParameters.OutputPath).Save();

            _performConversion = true;
            this.Close();

        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _performConversion = false;
            this.Close();
        }

        private void DetailedMetadataSection_Expanded(object sender, RoutedEventArgs e)
        {
            if (DocumentMetadata.DetailedMetadataSection.IsExpanded) {
                ConversionOptionsSection.IsExpanded = false;
            } else {
                ConversionOptionsSection.IsExpanded = true;
            }
        }
        private void ConversionOptionsSection_Expanded(object sender, RoutedEventArgs e)
        {
            if(ConversionOptionsSection.IsExpanded) {
                DocumentMetadata.DetailedMetadataSection.IsExpanded = false;
            } else {
                DocumentMetadata.DetailedMetadataSection.IsExpanded = true;
            }
        }
    }
}
