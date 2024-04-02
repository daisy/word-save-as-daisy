using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class BoolControl : BaseUserControl
    {
        BoolDataType m_boolDataType;

        private BoolControl() : base()
        {
            this.parameterNiceName.Text = "test";
            InitializeComponent();
        }

        public BoolControl(ScriptParameter p)
            : base(p)
        {
            InitializeComponent();
            setLinkedParameter(p);
        }

        public override void setLinkedParameter(ScriptParameter s)
        {
            base.setLinkedParameter(s);
            checkBox1.AccessibleName = s.NiceName;
            checkBox1.AccessibleDescription = s.Description;
            descriptionTooltip.SetToolTip(checkBox1, s.Description);
            m_boolDataType = (BoolDataType)s.ParameterDataType;
            checkBox1.Checked = (bool)s.ParameterValue;
        }



        public override void UpdateScriptParameterValue()
        {
            try
            {
                m_boolDataType.Value = checkBox1.Checked;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            this.UpdateScriptParameterValue();
        }
    }

}
