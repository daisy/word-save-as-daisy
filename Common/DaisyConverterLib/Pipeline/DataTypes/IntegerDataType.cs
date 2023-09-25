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
        

        public int Min { get; set; }
        public int Max { get; set; }

        public IntegerDataType(ScriptParameter p, XmlNode node) : base(p)
        {
            m_Parameter = p;
            PopulateListFromNode(node);
        }

        public IntegerDataType(int min = 0, int max = Int32.MaxValue, int defaultValue = Int32.MinValue) : base()
        {
            Min = min;
            Max = max;
            Value = defaultValue > Int32.MinValue ? defaultValue : min;
           
        }

        private void PopulateListFromNode(XmlNode DatatypeNode)
        {
            XmlNode EnumNode = DatatypeNode.FirstChild;
            Min = Convert.ToInt32(EnumNode.Attributes.GetNamedItem("min").Value);
            Max = Convert.ToInt32(EnumNode.Attributes.GetNamedItem("max").Value);
            Value = Min;
        }

        private int _value;
        public override object Value
        {
            get => _value;
            set
            {
                _value = (int)value;
                UpdateScript(value);
            }
        }


    }
}
