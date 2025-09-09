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
        private EnumData linkedEnumData;

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
            linkedEnumData = (EnumData)s.ParameterData;
            ValueSelector.AccessibleName = s.NiceName;
            ValueSelector.AccessibleDescription = s.Description;
            descriptionTooltip.SetToolTip(ValueSelector, s.Description);
            ValueSelector.Items.Clear();
            //base.DescriptionLabel = p.Description;
            EnumData enumData = (EnumData)s.ParameterData;
            foreach (string a in enumData.Keys)
            {
                ValueSelector.Items.Add(a);
            }
            base.Size = this.Size;
            if (linkedEnumData.SelectedIndex >= 0 && linkedEnumData.SelectedIndex < linkedEnumData.Values.Count)
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
