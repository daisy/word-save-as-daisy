using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion;


namespace Daisy.SaveAsDAISY.Forms.Controls {
    public partial class BoolControl : BaseUserControl
    {
        ScriptParameter m_Parameter;
        BoolDataType m_boolDataType;

        public BoolControl()
        {
            InitializeComponent();
        }

        public BoolControl(ScriptParameter p)
            : this()
        {
            // base.DescriptionLabel = p.Description;
            checkBox1.Text = p.NiceName;
            m_Parameter = p;
            m_boolDataType = (BoolDataType)p.ParameterDataType;
            checkBox1.Checked = m_boolDataType.Value;
            base.Size = this.Size;
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

        private void BoolControl_Load(object sender, EventArgs e)
        {

        }
    }

}
