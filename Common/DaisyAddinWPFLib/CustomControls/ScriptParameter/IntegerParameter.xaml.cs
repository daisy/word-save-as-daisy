using Daisy.SaveAsDAISY.Conversion;
using System.Text.RegularExpressions;
using System.Windows.Controls;
namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour IntegerParameter.xaml
    /// </summary>
    public partial class IntegerParameter : UserControl
    {
        public IntegerParameter()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        private readonly Regex IsNumber = new Regex(@"^\d*$");
        private void IntegerParameterTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = IsNumber.IsMatch(e.Text) || e.Text == string.Empty; // Allows only digits and empty input
        }

        private void IntegerParameterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ParameterValue = ((TextBox)sender).Text;
        }


        public ScriptParameter BoundParameter = null;

        public IntegerParameter(ScriptParameter p)
            : this()
        {
            BoundParameter = p;
        }

        public string ParameterName {
            get
            {
                return BoundParameter != null ? BoundParameter.NiceName : string.Empty;
            }
        }
        public string ParameterValue {
            get
            {
                return BoundParameter != null ? (string)BoundParameter.Value.ToString() : "0";
            }
            set
            {
                if (BoundParameter != null) {
                    BoundParameter.Value = int.Parse(value);
                }
            }
        }

        public string ParameterDesription {
            get
            {
                return BoundParameter != null ? (string)BoundParameter.Description.ToString() : "";
            }
        }
    }
}
