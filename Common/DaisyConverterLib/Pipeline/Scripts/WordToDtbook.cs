using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class WordToDtbook : WordBasedScript
    {
        public WordToDtbook(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-dtbook";
            this.niceName = "Export to DTBook XML";
            _parameters.Add("output", 
                new ScriptParameter(
                    "result",
                    "DTBook output",
                    new PathData(
                        PathData.InputOrOutput.output,
                        PathData.FileOrDirectory.Directory
                    ),
                    true,
                    "Output folder of the conversion to DTBook XML"
                )
            );
            _parameters.Add("narrator",
                new ScriptParameter(
                "narrator",
                "Prepare dtbook for pipeline 1 narrator",
                new BoolData(false),
                true,
                ""

                )
            );
            _parameters.Add("ApplySentenceDetection",
                new ScriptParameter(
                "ApplySentenceDetection",
                "Apply sentences detection",
                new BoolData(false),
                true,
                ""
                )
            );
           
        }
    }
}
