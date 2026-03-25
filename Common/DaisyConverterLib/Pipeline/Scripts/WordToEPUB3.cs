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
    public class WordToEPUB3 : WordBasedScript
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToEPUB3(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-epub3";
            this.niceName = "Export to EPUB3";
            _parameters.Add("output", 
                new ScriptParameter(
                    "result",
                    "EPUB3 output",
                    new PathData(
                        PathData.InputOrOutput.output,
                        PathData.FileOrDirectory.Directory
                    ),
                    true,
                    "Output folder of the conversion to EPUB3"
                )
            ); 
            _parameters.Add("validation", 
                new ScriptParameter(
                    "validation",
                    "Validation",
                    new EnumData(
                        new Dictionary<string, object> {
                            { "No validation", "off" },
                            { "Report validation issues", "report" },
                            { "Abort on validation issues", "abort" },
                        }, "off"),
                    false,
                    "Whether to abort on validation issues."
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
                    "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
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
            
        }
    }
}
