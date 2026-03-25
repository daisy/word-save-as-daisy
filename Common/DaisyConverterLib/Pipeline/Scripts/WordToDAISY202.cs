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
    public class WordToDAISY202 : WordBasedScript
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToDAISY202(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-daisy202";
            this.niceName = "Export to DAISY2.02";
            _parameters.Add("output",
                new ScriptParameter(
                    "result",
                    "DAISY2.02 output",
                    new PathData(
                        PathData.InputOrOutput.output,
                        PathData.FileOrDirectory.Directory
                    ),
                    true,
                    "Output folder of the conversion to DAISY2.02"
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
            //{ "validation-report", new ScriptParameter(
            //        "validation-report",
            //        "Validation reports",
            //        new PathData(
            //            PathData.InputOrOutput.output,
            //            PathData.FileOrDirectory.Directory
            //        ),
            //        false,
            //        "Output path of the validation reports",
            //        false,
            //        ParameterDirection.Output
            //    )
            //},
            _parameters.Add("language",
                new ScriptParameter(
                        "language",
                        "Language code",
                        new StringData(),
                        false,
                        "Language code of the input document."
                    )
                );
            _parameters.Add("tts-config", new ScriptParameter(
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
            _parameters.Add("audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolData(false),
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                );
        }
    }
}
