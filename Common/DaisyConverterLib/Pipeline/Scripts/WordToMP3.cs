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
    public class WordToMP3 : WordBasedScript
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToMP3(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-mp3";
            this.niceName = "Export to MP3";
            _alsoExportShapes = false;
            _parameters.Add("output", 
                new ScriptParameter(
                    "result",
                    "MP3 output",
                    new PathData(
                        PathData.InputOrOutput.output,
                        PathData.FileOrDirectory.Directory
                    ),
                    true,
                    "Output folder of the conversion to MP3"
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
            _parameters.Add("folder-depth",
                new ScriptParameter(
                "folder-depth",
                "Folder depth",
                new EnumData(
                    new Dictionary<string, object>(){
                        {"1", "1"},
                        {"2", "2"},
                        {"3", "3"},
                    }, "1"),
                false,
                "The number of folder levels in the produced folder structure.\n\n" +
                "The book is always, if possible, contained in a single top - level folder with MP3 files or\n" +
                "sub - folders(files for folder depth 1, sub - folders for folder depths greater than 1) that correspond\n" +
                "with top-level sections of the book.\n\n" +
                "If there are more top - level sections than the maximum number of files / folders that a top - level\n" +
                "folder can contain, the book is divided over multiple top-level folders.Similarly, if the number of\n" +
                "level - two sections within a top-level section exceeds the maximum number of files / folders that a\n" +
                "level - two folder can contain, the top-level section is divided over multiple level-two folders."

                )
            );
        }
    }
}
