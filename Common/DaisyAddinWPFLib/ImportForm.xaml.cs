using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.WPF.CustomControls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour ImportForm.xaml
    /// </summary>
    public partial class ImportForm : Window
    {
        public string ConversionTitle { get; set; }

        private bool _performConversion = false;

        public Script ScriptToRun { get; set; } = null;


        public ImportForm()
        {
            InitializeComponent();
        }

        public ImportForm(Script _script) : this()
        {
            if (_script != null)
            {
                PrepopulateDaisyOutput prepopulateDaisyOutput = PrepopulateDaisyOutput.Load();
                OutputDirectory.BoundParameter = _script.Parameters["output"];
                OutputDirectory.ParameterValue = prepopulateDaisyOutput != null
                                      ? prepopulateDaisyOutput.OutputPath
                                      : ConverterSettings.Instance.ResultsFolder;

                InputSelection.BoundParameter = _script.Parameters["input"];

                foreach (var item in _script.Parameters.Where(kv => kv.Key != "input" && kv.Key != "output"))
                {
                    ScriptParameter sp = item.Value;
                    ParameterDataBase parameter = sp.ParameterData;

                    if (sp.IsParameterRequired || sp.IsParameterDisplayed)
                    {
                        if (parameter is BoolData)
                        {
                            ScriptOptions.Children.Add(new CustomControls.BooleanParameter(sp));
                        }
                        else if (parameter is EnumData)
                        {
                            ScriptOptions.Children.Add(new CustomControls.EnumParameter(sp));
                        }
                        else if (parameter is StringData)
                        {
                            ScriptOptions.Children.Add(new CustomControls.StringParameter(sp));
                        }
                        else if (parameter is PathData)
                        {
                            ScriptOptions.Children.Add(new CustomControls.PathParameter(sp));
                        }
                        else if (parameter is IntegerData)
                        {
                            ScriptOptions.Children.Add(new CustomControls.IntegerParameter(sp));
                        }
                        else
                        {
                            ScriptOptions.Children.Add(new CustomControls.StringParameter(sp));
                        }
                    }
                }
                DataContext = this;
                ConversionTitle = _script.NiceName;
                ScriptToRun = _script;
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            bool formIsValid = true;
            string inputToCheck = InputSelection.ParameterValue as string;
            if (string.IsNullOrEmpty(inputToCheck) || !File.Exists(inputToCheck))
            {
                MessageBox.Show("Please select a valid input file", "Input file is required", MessageBoxButton.OK, MessageBoxImage.Warning);
                InputSelection.Focus();
                formIsValid = false;
            }
            string outputToCheck = OutputDirectory.ParameterValue as string;
            if (string.IsNullOrEmpty(outputToCheck) || !Directory.Exists(outputToCheck))
            {
                MessageBox.Show("Please select an existing output directory", "Output directory is required", MessageBoxButton.OK, MessageBoxImage.Warning);
                OutputDirectory.Focus();
                formIsValid = false;
            }

            foreach (UserControl item in ScriptOptions.Children)
            {
                
                ScriptParameter bound =
                    item is BooleanParameter b ? b.BoundParameter :
                    item is EnumParameter en ? en.BoundParameter :
                    item is StringParameter s ? s.BoundParameter :
                    item is PathParameter p ? p.BoundParameter :
                    item is IntegerParameter i ? i.BoundParameter :
                    null;
                if(bound != null && bound.IsParameterRequired && (string.IsNullOrEmpty(bound.Value?.ToString())))
                {
                    MessageBox.Show($"Please provide a valid value for {bound.Name}", "Invalid parameter value", MessageBoxButton.OK, MessageBoxImage.Warning);
                    item.Focus();
                    formIsValid = false;
                    break;
                }
            }

            if (formIsValid)
            {
                _performConversion = true;
                this.Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            _performConversion = false;
            this.Close();
        }

        public new bool ShowDialog()
        {
            base.ShowDialog();
            return _performConversion;
        }
    }
}
