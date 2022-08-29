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
        protected string m_Value;
        private string m_Regex;
        public StringDataType(ScriptParameter p, XmlNode DataTypeNode) : base(p)
        {
            m_Value = p.ParameterValue.ToString();
            XmlNode ChildNode = DataTypeNode.FirstChild;
            if (ChildNode.Attributes.Count > 0)
            {
                m_Regex = ChildNode.Attributes.GetNamedItem("regex").Value;
            }
        }

        public StringDataType(string defaultValue = "", string regex = "") : base() {
            m_Value = defaultValue;
            m_Regex = regex;
        }

        public override object Value
        {
            get { return m_Value; }
            set
            {
                if (value!=null)
                {
                    m_Value = (string)value;
                    UpdateScript(m_Value);
                }
                else throw new System.Exception("No_Title");
            }
        }

        private bool UpdateScript(string Val)
        {
            if (Val != null && m_Parameter != null)
            {
                m_Parameter.ParameterValue = Val;
                return true;
            }
            else
                return false;
        }
    }
}
