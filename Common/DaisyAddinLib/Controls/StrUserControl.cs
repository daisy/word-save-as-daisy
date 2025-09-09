using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class StrUserControl : BaseUserControl
    {
        private StringData linkedStringData;
        public StrUserControl()
        {
            InitializeComponent();
        }
        public StrUserControl(ScriptParameter p)
            : base(p)
        {
            InitializeComponent();
            setLinkedParameter(p);
        }

        public override void setLinkedParameter(ScriptParameter p)
        {
            base.setLinkedParameter(p);
            linkedStringData = (StringData)p.ParameterData;
            parameterInput.Text = p.Value.ToString();

            parameterInput.AccessibleName = p.NiceName;
            parameterInput.AccessibleDescription = p.Description;
            descriptionTooltip.SetToolTip(parameterInput, p.Description);
        }

        public override void UpdateScriptParameterValue()
        {
            try
            {
                linkedStringData.Value = parameterInput.Text;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
