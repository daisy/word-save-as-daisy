using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{

    public class ScriptParameter
    {

        public enum ParameterDirection
        {
            Input,
            Output,
            Option
        }

        private string m_Name;
        private object m_Value;
        private string m_NiceName;
        private bool m_Required;
        private string m_Description;
        private ParameterDataType m_DataType;

        private bool m_Displayed;

        private ParameterDirection m_Direction;


        public ScriptParameter(
            string name,
            string niceName,
            ParameterDataType dataType,
            object initialValue,
            bool required = false,
            string description = "",
            bool displayed = true,
            ParameterDirection directon = ParameterDirection.Option
        )
        {
            m_Name = name;
            m_NiceName = niceName;
            m_DataType = dataType;
            m_DataType.LinkedParameter = this;
            m_Value = initialValue;
            m_Description = description;
            m_Required = required;
            m_Displayed = displayed;
            m_Direction = directon;
        }

        public ScriptParameter(XmlNode node)
        {
            GetParameterAttributeInfo(node);
            GetParameterProperties(node);
        }


        /// <summary>
        /// extracts  parameter attribute info from script xml file
        /// </summary>
        /// <param name="node"></param>
        private void GetParameterAttributeInfo(XmlNode node)
        {
            // Get attribute  information of ScriptParameter
            for (int AttrIndex = 0; AttrIndex < node.Attributes.Count; AttrIndex++)
            {

                switch (node.Attributes[AttrIndex].Name)
                {
                    case "name":
                        m_Name = node.Attributes.GetNamedItem("name").Value;
                        break;

                    case "value":
                        m_Value = node.Attributes.GetNamedItem("value").Value;
                        break;

                    case "required":
                        m_Required = node.Attributes.GetNamedItem("required").Value == "true" ? true : false;
                        break;
                }
            }
        }

        /// <summary>
        ///  extracts parameter properties from script xml file
        /// </summary>
        /// <param name="node"></param>
        private void GetParameterProperties(XmlNode node)
        {
            // Get properties of parameter from its child nodes
            XmlNode ChildNode = node.FirstChild;

            while (ChildNode != null)
            {
                switch (ChildNode.Name)
                {
                    case "nicename":
                        m_NiceName = ChildNode.InnerText;
                        break;

                    case "description":
                        m_Description = ChildNode.InnerText;
                        break;

                    case "datatype":
                        GetDatatype(ChildNode);
                        break;
                }

                ChildNode = ChildNode.NextSibling;
            }

        }


        /// <summary>
        /// choose  datatype matching  parameter datatype in script and set its properties from script xml file
        /// </summary>
        /// <param name="DatatypeNode"></param>
        private void GetDatatype(XmlNode DatatypeNode)
        {
            switch (DatatypeNode.FirstChild.Name)
            {
                case "file":
                case "directory":
                    m_DataType = new PathDataType(this, DatatypeNode);
                    break;

                case "enum":
                    m_DataType = new EnumDataType(this, DatatypeNode);
                    break;

                case "boolean":
                    m_DataType = new BoolDataType(this, DatatypeNode);
                    break;

                case "integer":
                    m_DataType = new IntegerDataType(this, DatatypeNode);
                    break;

                case "string":
                    m_DataType = new StringDataType(this, DatatypeNode);
                    break;
            }
        }


        public string Name { get { return m_Name; } }
        public string NiceName { get { return m_NiceName; } }
        public string Description { get { return m_Description; } }
        public bool IsParameterRequired { get { return m_Required; } }

        public bool IsParameterDisplayed { get { return m_Displayed; } }

        /// <summary>
        /// Direction of the parameter (input, output or option)
        /// </summary>
        public ParameterDirection Direction { get { return m_Direction; } }

        public ParameterDataType ParameterDataType { get { return m_DataType; } }

        public object ParameterValue
        {
            get { return m_Value; }
            set
            {
                if (value != null)
                {
                    m_Value = value;
                    m_Required = true;
                }
                else if (m_Required && (m_Value == null))
                    m_Required = false;
            }
        }

    }

}

