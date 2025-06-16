using System;
using System.Collections;
using System.Collections.Generic;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class StringData : ParameterDataBase
    {
        
        private string m_Regex;

        public StringData(string defaultValue = "", string regex = "") : base(defaultValue)
        {
            m_Regex = regex;
        }
    }
}
