using Daisy.SaveAsDAISY.Conversion;
using System.Windows.Controls;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour BooleanParameter.xaml
    /// </summary>
    /// 
    public partial class BooleanParameter : UserControl
    {
        public BooleanParameter()
        {
            InitializeComponent();
            this.DataContext = this;
        }


        public ScriptParameter BoundParameter = null;

        public BooleanParameter(ScriptParameter p)
            : this()
        {
            BoundParameter = p;
        }

        public string ParameterName { get {
                return BoundParameter != null ? BoundParameter.NiceName : string.Empty;
            }
        }
        public bool ParameterValue {
            get
            {
                return BoundParameter != null ? (bool)BoundParameter.Value : false;
            }
            set {
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
