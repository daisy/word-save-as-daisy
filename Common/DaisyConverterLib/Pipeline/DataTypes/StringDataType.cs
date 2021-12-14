using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class StringDataType : ParameterDataType {
        private string m_Title;
        private string m_strValue;
        public StringDataType(ScriptParameter p, XmlNode DataTypeNode) : base(p)
        {
            m_Title = p.ParameterValue;
            XmlNode ChildNode = DataTypeNode.FirstChild;
            if (ChildNode.Attributes.Count > 0)
            {
                m_strValue = ChildNode.Attributes.GetNamedItem("regex").Value;
            }
        }

        public StringDataType(string title, string regex = "") : base() {
            m_Title = title;
            m_strValue = regex;
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
