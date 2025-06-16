using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class DtbookToDaisy3 : Pipeline2Script
    {
        public DtbookToDaisy3(IConversionEventsHandler eventsHandler) : base(eventsHandler)
        {
            this.name = "dtbook-to-daisy3";
            _parameters = new Dictionary<string, ScriptParameter> {
                {"input", new ScriptParameter(
                        "source",
                        "input",
                        new PathData(PathData.InputOrOutput.input,PathData.FileOrDirectory.File),
                        true,
                        "input",
                        false,
                        ParameterDirection.Input
                    )
                },
                //{"validation-status", new ScriptParameter(
                //    "validation-status",
                //    "validation-status"
                //    )
                //},
                {"include-tts-log", new ScriptParameter(
                        "include-tts-log",
                        "include-tts-log",
                        new BoolData(false),
                        false,
                        "Include tts log with the result",
                        false
                    )
                },
                {"tts-log", new ScriptParameter(
                        "tts-log",
                        "TTS log output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        false,
                        "TTS log output directory",
                         true,
                        ParameterDirection.Output
                    )
                },
                {"publisher", new ScriptParameter(
                        "publisher",
                        "Publisher",
                        new StringData(""),
                        false,
                        "The agency responsible for making the Digital Talking Book available. If left blank, it will be retrieved from the DTBook meta-data.",
                        false
                    )
                },
                {"output", new ScriptParameter(
                        "output-dir",
                        "Daisy3 output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "The resulting DAISY 3 publication."
                    )
                },
                {"tts-config", new ScriptParameter(
                        "tts-config",
                        "Text-to-speech configuration file",
                        new PathData(PathData.InputOrOutput.input,PathData.FileOrDirectory.File),
                        false,
                        "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                        true,
                        ParameterDirection.Input
                    )
                },
                {"audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolData(true),
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
                {"with-text", new ScriptParameter(
                        "with-text",
                        "With text",
                        new BoolData(true),
                        false,
                        "Includes DTBook in output, as opposed to audio only."
                    )
                },
                {"word-detection", new ScriptParameter(
                        "word-detection",
                        "Apply word detection",
                        new BoolData(false),
                        false,
                        "Whether to detect and mark up words with <w> tags.\r\n" +
                        "\r\n" +
                        "By default word detection is performed but an option is provided to disable it because some DAISY 3\r\n" +
                        "reading systems can't handle the word tags."
                    )
                },
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputDirectory"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">if the file was not found</exception>
        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            return SearchInputFromDirectory(inputDirectory);
        }

        public static string SearchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            return Directory.GetFiles(inputDirectory.FullName, "*.xml", SearchOption.AllDirectories)[0];
        }
    }
}
