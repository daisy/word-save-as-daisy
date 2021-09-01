using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class StringDataType
    {
        public ScriptParameter m_Parameter;
        private string m_Title;
        private string m_strValue;
        public StringDataType(ScriptParameter p, XmlNode DataTypeNode)
        {
            m_Parameter = p;
            m_Title = p.ParameterValue;
            XmlNode ChildNode = DataTypeNode.FirstChild;
            if (ChildNode.Attributes.Count > 0)
            {
                m_strValue = ChildNode.Attributes.GetNamedItem("regex").Value;
            }
        }

        public string Value
        {
            get { return m_Title; }
            set
            {
                if (value!=null)
                {
                    m_Title = value;
                    UpdateScript(m_Title);
                }
                else throw new System.Exception("No_Title");
            }
        }

        private bool UpdateScript(string Val)
        {
            if (Val != null)
            {
                m_Parameter.ParameterValue = Val;
                return true;
            }
            else
                return false;
        }
    }
}
