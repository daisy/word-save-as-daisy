using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms.Controls {
    public partial class StrUserControl : BaseUserControl
    {
        private ScriptParameter m_Parameter;
        private StringDataType m_StrData;
        public StrUserControl()
        {
            InitializeComponent();
        }
        public StrUserControl(ScriptParameter p)
            : this()
        {
            m_Parameter = p;
            m_StrData = (StringDataType)p.ParameterDataType;
            int wdiff = m_StrNiceLbl.Width;
            m_StrNiceLbl.Text = p.NiceName;
            StringDataType StrData = (StringDataType)p.ParameterDataType;
            wdiff -= m_StrNiceLbl.Width;
            if (wdiff < 0)
            {
                Point location = m_StrNiceLbl.Location;
                Width -= wdiff;
                m_StrNiceLbl.Location = location;
            }
            else
            {
                m_StrNiceLbl.Location = new Point(m_StrNiceLbl.Location.X - wdiff, m_StrNiceLbl.Location.Y);
            }
        }

        public override void UpdateScriptParameterValue()
        {
            try
            {
                m_StrData.Value = m_StrTxtBox.Text;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
