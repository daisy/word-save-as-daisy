using Daisy.SaveAsDAISY.Conversion;
using System.Windows.Controls;
namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour StringParameter.xaml
    /// </summary>
    public partial class StringParameter : UserControl
    {
        public StringParameter()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        private ScriptParameter BoundParameter = null;

        public StringParameter(ScriptParameter p)
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
                return BoundParameter != null ? BoundParameter.Value.ToString() : "";
            }
            set
            {
                if (BoundParameter != null) {
                    BoundParameter.Value = value;
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
