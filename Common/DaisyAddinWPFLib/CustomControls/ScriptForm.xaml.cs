using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour ScriptForm.xaml
    /// </summary>
    public partial class ScriptForm : StackPanel
    {
        public ScriptForm(Script script, List<string> hiddenFieldsKeys = null)
        {
            InitializeComponent();
            List<string> hiddenFieldsKeysCombined = 
                (hiddenFieldsKeys ?? new List<string>())
                .Concat(new List<string> { "input", "output" }).ToList();
            foreach (var item in script.Parameters.Where(kv => !hiddenFieldsKeysCombined.Contains(kv.Key)))
            {
                ScriptParameter sp = item.Value;
                ParameterDataBase parameter = sp.ParameterData;

                if (sp.IsParameterRequired || sp.IsParameterDisplayed)
                {
                    if (parameter is BoolData)
                    {
                        this.Children.Add(new CustomControls.BooleanParameter(sp));
                    }
                    else if (parameter is EnumData)
                    {
                        this.Children.Add(new CustomControls.EnumParameter(sp));
                    }
                    else if (parameter is StringData)
                    {
                        this.Children.Add(new CustomControls.StringParameter(sp));
                    }
                    else if (parameter is PathData)
                    {
                        this.Children.Add(new CustomControls.PathParameter(sp));
                    }
                    else if (parameter is IntegerData)
                    {
                        this.Children.Add(new CustomControls.IntegerParameter(sp));
                    }
                    else
                    {
                        this.Children.Add(new CustomControls.StringParameter(sp));
                    }
                }
            }
        }
    }
}
