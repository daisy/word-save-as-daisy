using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Daisy.SaveAsDAISY.Conversion.Pipeline;

using Newtonsoft.Json;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class MapDataType : StringDataType
    {
        public override object Value
        {
            get => (object)JsonConvert.DeserializeObject<Dictionary<string, object>>(this.m_Value);
            set => this.Value = value;
        }
    }
}
