using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Daisy.SaveAsDAISY.Conversion;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class IntUserControl : BaseUserControl
    {
        IntegerDataType m_IntDataType;

        public IntUserControl()
        {
            InitializeComponent();
        }
        public IntUserControl(ScriptParameter p)
            : base(p)
        {
            InitializeComponent();
            setLinkedParameter(p);
        }

        public override void setLinkedParameter(ScriptParameter p)
        {
            base.setLinkedParameter(p);
            m_IntDataType = (IntegerDataType)p.ParameterDataType;
            mListBokBox.AccessibleName = p.NiceName;
            mListBokBox.AccessibleDescription = p.Description;
            descriptionTooltip.SetToolTip(mListBokBox, p.Description);

            IntegerDataType IntData = (IntegerDataType)p.ParameterDataType;

            mListBokBox.Minimum = IntData.Min;
            mListBokBox.Maximum = IntData.Max;
            mListBokBox.Value = (int)IntData.Value;

            base.Size = this.Size;
        }


        public override void UpdateScriptParameterValue()
        {
            try
            {
                m_IntDataType.Value = Convert.ToInt32(mListBokBox.Value);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


    }
}
