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

            List<string> intList = IntData.GetValues;
            mListBokBox.Minimum = Convert.ToInt32(intList[0]);
            mListBokBox.Maximum = Convert.ToInt32(intList[1]);
            base.Size = this.Size;
        }


        public override void UpdateScriptParameterValue()
        {
            try
            {
                m_IntDataType.SelectedIndex = Convert.ToInt32(mListBokBox.Value);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


    }
}
