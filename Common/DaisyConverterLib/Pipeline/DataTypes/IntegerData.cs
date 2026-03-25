using System;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class IntegerData : ParameterDataBase
    {
        public int Min { get; set; }
        public int Max { get; set; }

        public IntegerData(int min = 0, int max = Int32.MaxValue, int defaultValue = Int32.MinValue) : base()
        {
            Min = min;
            Max = max;
            Value = defaultValue > Int32.MinValue ? defaultValue : min;
           
        }



    }
}
