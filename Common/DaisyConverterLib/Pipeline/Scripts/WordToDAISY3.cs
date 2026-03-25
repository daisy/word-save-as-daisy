using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class WordToDAISY3 : WordBasedScript
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToDAISY3(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-daisy3";
            this.niceName = "Export to DAISY3";
            _parameters.Add("output",
                new ScriptParameter(
                    "result",
                    "DAISY3 output",
                    new PathData(
                        PathData.InputOrOutput.output,
                        PathData.FileOrDirectory.Directory
                    ),
                    true,
                    "Output folder of the conversion to DAISY3"
                )
            );
            _parameters.Add("include-tts-log",
                new ScriptParameter(
                    "include-tts-log",
                    "include-tts-log",
                    new BoolData(false),
                    false,
                    "Include tts log with the result",
                    false
                )
            );
            _parameters.Add("tts-config",
                new ScriptParameter(
                    "tts-config",
                    "Text-to-speech configuration file",
                    new PathData(
                        PathData.InputOrOutput.input,
                        PathData.FileOrDirectory.File,
                        "",
                        GlobaleSettings.TTSConfigFile ?? ""
                    ),
                    false,
                    "Configuration file for the text-to-speech.\r\n\r\n" +
                    "[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                    !GlobaleSettings.UseDAISYPipelineApp,
                    ParameterDirection.Input
                )
            ); 
            _parameters.Add("audio",
                new ScriptParameter(
                    "audio",
                    "Enable text-to-speech",
                    new BoolData(true),
                    false,
                    "Whether to use a speech synthesizer to produce audio files."
                )
            );
            _parameters.Add("with-text", 
                new ScriptParameter(
                    "with-text",
                    "With text",
                    new BoolData(true),
                    false,
                    "Includes DTBook in output, as opposed to audio only."
                )
            );
            _parameters.Add("word-detection", 
                new ScriptParameter(
                    "word-detection",
                    "Apply word detection",
                    new BoolData(false),
                    false,
                    "Whether to detect and mark up words with <w> tags.\r\n" +
                    "\r\n" +
                    "By default word detection is performed but an option is provided to disable it because some DAISY 3\r\n" +
                    "reading systems can't handle the word tags."
                )
            );
            
        }
    }
}
