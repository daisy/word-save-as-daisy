using Daisy.SaveAsDAISY.Conversion;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour EnumParameter.xaml
    /// </summary>
    public partial class EnumParameter : UserControl
    {
        public EnumParameter()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public ScriptParameter BoundParameter = null;

        private Dictionary<string, object> _parameterValues = new Dictionary<string, object>();

        public EnumParameter(ScriptParameter p)
            : this()
        {
            BoundParameter = p;
            if (p.ParameterData is EnumData enumData)
            {
                for (int i = 0; i < enumData.Values.Count; i++) {
                    _parameterValues.Add(enumData.Keys[i], enumData.Values[i]);
                }
            }
            
        }

        public string ParameterName {
            get
            {
                return BoundParameter != null ? BoundParameter.NiceName : string.Empty;
            }
        }

        public object ParameterValue {
            get
            {
                return BoundParameter != null ? BoundParameter.Value : "";
            }
            set
            {
                if (BoundParameter != null) {
                    BoundParameter.Value = value;
                }
            }
        }

        public Dictionary<string, object> ParameterValues {
            get
            {
                return _parameterValues;
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
