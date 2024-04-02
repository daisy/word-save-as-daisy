﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    abstract public class ParameterDataType
    {

        protected ScriptParameter m_Parameter;

        public ScriptParameter LinkedParameter { get => m_Parameter; set
            {
                // Reference the script parameter
                m_Parameter = value;
                // Get its value
                Value = m_Parameter.ParameterValue;
            }
        }

        public ParameterDataType(ScriptParameter param)
        {
            this.m_Parameter = param;
        }

        public ParameterDataType()
        {
        }

        public abstract object Value { get; set; }

        protected bool UpdateScript(object Val)
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
