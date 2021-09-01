using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms.Controls {
    public partial class IntUserControl : BaseUserControl
    {
        ScriptParameter m_Parameter;
        IntegerDataType m_IntDataType;
        
        public IntUserControl()
        {
            InitializeComponent();
        }
        public IntUserControl(ScriptParameter p)
            : this()
        {
            mIntNiceLabel.Text = p.NiceName;
            m_Parameter = p;
            m_IntDataType = (IntegerDataType)p.ParameterDataType;
            int wdiff = mIntNiceLabel.Width;
            IntegerDataType IntData = (IntegerDataType)p.ParameterDataType;

            wdiff -= mIntNiceLabel.Width;
            if (wdiff < 0)
            {
                Point location = mIntNiceLabel.Location;
                Width -= wdiff;
                mIntNiceLabel.Location = location;
            }
            else
            {
                mIntNiceLabel.Location = new Point(mIntNiceLabel.Location.X - wdiff, mIntNiceLabel.Location.Y);
            }
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
