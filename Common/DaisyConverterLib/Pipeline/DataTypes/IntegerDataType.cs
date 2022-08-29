using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class IntegerDataType : ParameterDataType
    {
        private List<int> m_ValueList;
        private List<string> m_NiceNameList;
        private int m_SelectedIndex;

        public IntegerDataType(ScriptParameter p, XmlNode node) : base(p) {
            m_Parameter = p;
            m_ValueList = new List<int>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            PopulateListFromNode(node);
        }

        public IntegerDataType(int min, int max) : base()
        {
            m_ValueList = new List<int>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            for (int i = min; i <= max; i++) {
                m_ValueList.Add(i);
                m_NiceNameList.Add(i.ToString());
            }
        }

        public IntegerDataType(List<int> possibleValues = null) : base()
        {
            m_ValueList = new List<int>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = 0;
            if(possibleValues != null)
            {
                foreach (int value in possibleValues)
                {
                    m_ValueList.Add(value);
                    m_NiceNameList.Add(value.ToString());

                }
            }
        }

        private void PopulateListFromNode(XmlNode DatatypeNode)
        {
            XmlNode EnumNode = DatatypeNode.FirstChild;
            m_ValueList.Add(Convert.ToInt32(EnumNode.Attributes.GetNamedItem("min").Value));
            m_ValueList.Add(Convert.ToInt32(EnumNode.Attributes.GetNamedItem("max").Value));
            for (int i = Convert.ToInt32(m_ValueList[0]); i <= Convert.ToInt32(m_ValueList[1]); i++)
            {

                m_NiceNameList.Add(i.ToString());
            }

        }

        public int SelectedIndex
        {
            get { return Convert.ToInt32(m_SelectedIndex); }
            set
            {
                if (m_NiceNameList.Contains(value.ToString()))
                    m_SelectedIndex = value;
                else throw new System.Exception("IndexNotInRange");
            }
        }
        public List<string> GetValues { get { return m_NiceNameList; } }

        public override object Value { 
            get => m_ValueList[SelectedIndex]; 
            set {
                SelectedIndex = m_ValueList.IndexOf((int)value);
            } 
        }
        

    }
}
