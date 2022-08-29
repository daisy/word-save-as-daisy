using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class EnumDataType : ParameterDataType {
        private List<string> m_ValueList;
        private List<string> m_NiceNameList;
        private int m_SelectedIndex;

        public EnumDataType(ScriptParameter p, XmlNode node) : base(p)
        {
            m_ValueList = new List<string>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            PopulateListFromNode(node);

            // select list node as default value if default  exists
            if (p.ParameterValue != null && p.ParameterValue.ToString() != ""
                && m_ValueList.Contains(p.ParameterValue.ToString()))
                m_SelectedIndex = m_ValueList.BinarySearch(p.ParameterValue.ToString());
        }

        public EnumDataType(Dictionary<string,string> itemsList, string defaultValue = "" ) : base() {
            m_ValueList = new List<string>();
            m_NiceNameList = new List<string>();
            m_SelectedIndex = -1;
            foreach(KeyValuePair<string,string> item in itemsList) {
                m_NiceNameList.Add(item.Key);
                m_ValueList.Add(item.Value);
            }
            
            if (defaultValue != null && defaultValue != ""
                && m_ValueList.Contains(defaultValue))
                m_SelectedIndex = m_ValueList.BinarySearch(defaultValue);
        }

        private void PopulateListFromNode(XmlNode DatatypeNode)
        {
            XmlNode EnumNode = DatatypeNode.FirstChild;
            XmlNodeList nodeList = EnumNode.SelectNodes("item");

            foreach (XmlNode n in nodeList)
            {
                m_ValueList.Add(n.Attributes.GetNamedItem("value").Value);
                m_NiceNameList.Add(n.Attributes.GetNamedItem("nicename").Value);
            }
        }


        public List<string> GetValues { get { return m_ValueList; } }
        public List<string> GetNiceNames { get { return m_NiceNameList; } }

        /// <summary>
        ///  Gets and sets the index of value selected.
        /// </summary>
        public int SelectedIndex
        {
            get { return m_SelectedIndex; }
            set
            {
                if (value >= 0 && value < m_ValueList.Count)
                    SetSelectedIndexAndUpdateScript(value);
                else throw new System.Exception("IndexNotInRange");
            }
        }


        public string SelectedItemValue
        {
            get { return m_ValueList[m_SelectedIndex]; }
            set
            {
                if (value != null && m_ValueList.Contains(value))
                    SetSelectedIndexAndUpdateScript(m_ValueList.BinarySearch(value));
                else throw new System.Exception("NotAbleToSelectItem");
            }
        }


        public string SelectedItemNiceName
        {
            get { return m_NiceNameList[m_SelectedIndex]; }
            set
            {
                if (value != null && m_NiceNameList.Contains(value))
                    SetSelectedIndexAndUpdateScript(m_NiceNameList.BinarySearch(value));
                else throw new System.Exception("NotAbleToSelectItem");
            }
        }

        public override object Value { 
            get => SelectedItemValue; 
            set {
                SelectedItemNiceName = value.ToString();
            } 
        }

        private bool SetSelectedIndexAndUpdateScript(int Index)
        {
            if (Index > 0 && Index < m_ValueList.Count)
            {
                m_SelectedIndex = Index;
                UpdateScript(SelectedItemValue);
                return true;
            }
            else
                return false;
        }
    }
}
