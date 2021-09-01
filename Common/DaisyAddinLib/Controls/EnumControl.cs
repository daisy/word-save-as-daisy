using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms.Controls {
    public partial class EnumControl : BaseUserControl
    {
        private ScriptParameter m_Parameter;
        private EnumDataType m_EnumData;

        public EnumControl()
        {
            InitializeComponent();
        }

        public EnumControl(ScriptParameter p)
            : this()
        {
            m_Parameter = p;
            m_EnumData = (EnumDataType)p.ParameterDataType;
            int wdiff = mNiceNameLabel.Width;
            mNiceNameLabel.Text = p.NiceName;
            //Point a = (Point) mNiceNameLabel.Width;
            mComboBox.AccessibleName = p.NiceName;
            mComboBox.Items.Clear();
            //base.DescriptionLabel = p.Description;
            EnumDataType EnumData = (EnumDataType)p.ParameterDataType;            
            wdiff -= mNiceNameLabel.Width;
            if (wdiff < 0)
            {

                Point location = mNiceNameLabel.Location;
                //mComboBox.Location = new Point(mNiceNameLabel.Location.X + 42, mComboBox.Location.Y);
                Width -= wdiff;
                mNiceNameLabel.Location = location;
            }
            else
            {
                mNiceNameLabel.Location = new Point(mNiceNameLabel.Location.X - wdiff, mNiceNameLabel.Location.Y);
            }
            foreach (string s in EnumData.GetNiceNames)
            {
                mComboBox.Items.Add(s);                
            }
            base.Size = this.Size;
            if (m_EnumData.SelectedIndex >= 0 && m_EnumData.SelectedIndex < m_EnumData.GetValues.Count)
                mComboBox.SelectedIndex = m_EnumData.SelectedIndex;
        }


        private void ComboboxControl_Load(object sender, EventArgs e)
        {

        }

        public override void UpdateScriptParameterValue()
        {
            try
            {
                m_EnumData.SelectedIndex = mComboBox.SelectedIndex;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void mLabel_Click(object sender, EventArgs e)
        {

        }

        private void mComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void mComboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }


    }

}
