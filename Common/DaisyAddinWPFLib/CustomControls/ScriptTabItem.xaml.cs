
using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using System.Linq;
using System.Windows.Automation;
using System.Windows.Controls;

namespace Daisy.SaveAsDAISY.WPF.CustomControls
{
    /// <summary>
    /// Logique d'interaction pour ScriptTabItem.xaml
    /// </summary>
    public partial class ScriptTabItem : TabItem
    {
        public Script Script { get; private set; }
        public ScriptTabItem(Script script)
        {
            InitializeComponent();
            Script = script;
            Header = Script.NiceName;
            AutomationProperties.SetHelpText(this, Script.Description);
            Content = new ScriptForm(Script);
        }
    }
}
