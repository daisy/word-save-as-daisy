using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    abstract public class ParameterDataBase
    {

        public ParameterDataBase(object value)
        {
            Value = value;
        }

        public ParameterDataBase()
        {
        }

        public virtual object Value { get; set; }


    }
}
