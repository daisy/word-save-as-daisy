
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
            // TODO : for now we hide the default settings keys here, this might need a change
            // Only language is hidden, as it is computed from the document properties
            Content = new ScriptForm(Script, new System.Collections.Generic.List<string>() { "language" });
        }
    }
}
