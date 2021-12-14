using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline {
    abstract public class ParameterDataType {

        protected ScriptParameter m_Parameter;

        public ScriptParameter LinkedParameter { get => m_Parameter; set => m_Parameter = value; }

        public ParameterDataType(ScriptParameter param) {
            this.m_Parameter = param;
        }

        public ParameterDataType() {
        }

    }
}
