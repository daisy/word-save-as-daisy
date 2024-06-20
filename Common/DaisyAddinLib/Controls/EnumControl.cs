using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class EnumControl : BaseUserControl
    {
        private EnumDataType linkedEnumData;

        public EnumControl()
        {
            InitializeComponent();
        }

        public EnumControl(ScriptParameter p)
            : base(p)
        {
            InitializeComponent();
            this.setLinkedParameter(p);
        }

        public override void setLinkedParameter(ScriptParameter s)
        {
            base.setLinkedParameter(s);
            linkedEnumData = (EnumDataType)s.ParameterDataType;
            ValueSelector.AccessibleName = s.NiceName;
            ValueSelector.AccessibleDescription = s.Description;
            descriptionTooltip.SetToolTip(ValueSelector, s.Description);
            ValueSelector.Items.Clear();
            //base.DescriptionLabel = p.Description;
            EnumDataType EnumData = (EnumDataType)s.ParameterDataType;
            foreach (string a in EnumData.GetNiceNames)
            {
                ValueSelector.Items.Add(a);
            }
            base.Size = this.Size;
            if (linkedEnumData.SelectedIndex >= 0 && linkedEnumData.SelectedIndex < linkedEnumData.GetValues.Count)
                ValueSelector.SelectedIndex = linkedEnumData.SelectedIndex;
        }

        public override void UpdateScriptParameterValue()
        {
            try
            {
                linkedEnumData.SelectedIndex = ValueSelector.SelectedIndex;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ValueSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateScriptParameterValue();
        }
    }

}
