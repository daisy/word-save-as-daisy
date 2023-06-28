using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class DtbookToEpub3 : Pipeline2Script
    {
        public DtbookToEpub3(IConversionEventsHandler eventsHandler) : base(eventsHandler)
        {
            this.name = "dtbook-to-epub3";
            _parameters = new Dictionary<string, ScriptParameter>
            {
                { "input", new ScriptParameter(
                        "source",
                        "OPF",
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File,
                            "application/oebps-package+xml"
                        ),
                        "",
                        true,
                        "The package file of the input DTB.",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                {"tts-log", new ScriptParameter(
                        "tts-log",
                        "TTS log output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        false,
                        "TTS log output directory",
                         true,
                        ScriptParameter.ParameterDirection.Output
                    )
                },
                {"language", new ScriptParameter(
                        "language",
                        "Language code",
                        new StringDataType(),
                        "",
                        false,
                        "Language code of the input document."
                    )
                },
                {"output", new ScriptParameter(
                        "output-dir",
                        "EPUB output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The produced EPUB."
                    )
                },
                { "validation", new ScriptParameter(
                        "validation",
                        "Validation",
                        new EnumDataType(
                            new Dictionary<string, object> {
                                { "No validation", "off" },
                                { "Report validation issues", "report" },
                                { "Abort on validation issues", "abort" },
                            }, "Abort on validation issues"),
                        "abort",
                        false,
                        "Whether to abort on validation issues."
                    )
                },
                { "validation-report", new ScriptParameter(
                        "validation-report",
                        "Validation reports",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        false,
                        "Output path of the validation reports",
                        false,
                        ScriptParameter.ParameterDirection.Output
                    )
                },
                //{ "validation-status", new ScriptParameter(
                //        "validation-report",
                //        "Validation reports",
                //        new PathDataType(
                //            PathDataType.InputOrOutput.output,
                //            PathDataType.FileOrDirectory.Directory
                //        ),
                //        "",
                //        false,
                //        ""
                //    )
                //},
                {"nimas", new ScriptParameter(
                        "nimas",
                        "NIMAS input",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether the input DTBook is a NIMAS 1.1-conformant XML content file.",
                        false // not sure this option should available in save as daisy
                    )
                },
                {"tts-config", new ScriptParameter(
                        "tts-config",
                        "Text-to-speech configuration file",
                        new PathDataType(PathDataType.InputOrOutput.input,PathDataType.FileOrDirectory.File),
                        "",
                        false,
                        "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                {"audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
            };
        }
    }
}
