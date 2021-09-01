using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion
{
    public enum ConversionStatus
    {
        None,
        Canceled,
        Error,
        ValidatedName,
        CreatedWorkingCopy,
        ProcessedShapes,
        ProcessedMathML,
        PreprocessingSucceeded,
        ReadyForConversion,
        HasStartedConversion,
        ConversionToDTBookSucceeded,
        PostProcessingSucceeded
    }
}
