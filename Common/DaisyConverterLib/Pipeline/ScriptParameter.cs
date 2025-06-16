using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using System.Runtime.Serialization;

namespace Daisy.SaveAsDAISY.Conversion
{
    public enum ParameterDirection
    {
        Input,
        Output,
        Option
    }

    

    public class ScriptParameter
    {
        public ScriptParameter() { }

        public ParameterDataBase ParameterData { get; private set; }

        public ScriptParameter(
            string name,
            string niceName,
            ParameterDataBase defaultValue,
            bool required = false,
            string description = "",
            bool displayed = true,
            ParameterDirection directon = ParameterDirection.Option
        )
        {
            Name = name;
            NiceName = niceName;
            ParameterData = defaultValue;
            Description = description;
            IsParameterRequired = required;
            IsParameterDisplayed = displayed;
            Direction = directon;
        }


        public string Name { get; private set; }
        public string NiceName { get; private set; }
        public string Description { get; private set; }
        public bool IsParameterRequired { get; private set; }
        public bool IsParameterDisplayed { get; private set; }


        /// <summary>
        /// Direction of the parameter (input, output or option)
        /// </summary>
        public ParameterDirection Direction { get; private set; }

        public object Value
        {
            get { return ParameterData.Value; }
            set
            {
                if (value != null)
                {
                    ParameterData.Value = value;
                    IsParameterRequired = true;
                }
                else if (IsParameterRequired && (ParameterData == null))
                    IsParameterRequired = false;
            }
        }
    }

}

