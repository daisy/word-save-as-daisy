using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class WordToMP3 : Pipeline2Script
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToMP3(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-mp3";
            this.niceName = "Export to MP3";
            _parameters = new Dictionary<string, ScriptParameter>
            {
                {
                    "input", new ScriptParameter(
                        "source",
                        "Input Docx file",
                        new PathData(
                            PathData.InputOrOutput.input,
                            PathData.FileOrDirectory.File
                        ),
                        true,
                        "The document you want to convert."
                    )
                },
                {
                    "output", new ScriptParameter(
                        "result",
                        "MP3 output",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "Output folder of the conversion to MP3"
                    )
                },
                { "title",
                    new ScriptParameter(
                        "title",
                        "Document title",
                        new StringData(),
                        false,"",false
                    )
                },
                { "creator",
                    new ScriptParameter(
                        "creator",
                        "Document creator or author",
                        new StringData(),
                        false,"",false
                    )
                },
                { "publisher",
                    new ScriptParameter(
                        "publisher",
                        "Document publisher",
                        new StringData(),
                        false,"",false
                    )
                },
                { "uid",
                    new ScriptParameter(
                        "uid",
                        "Document identifier",
                        new StringData(),
                        false,
                        "Identifier to be added as dtb:uid metadata",
                        false
                    )
                },
                { "subject",
                    new ScriptParameter(
                        "subject",
                        "Subject(s)",
                        new StringData(),
                        false,
                        "Subject(s) to be added as dc:Subject metadata",
                        false
                    )
                },
                { "accept-revisions",
                    new ScriptParameter(
                        "accept-revisions",
                        "Accept revisions",
                        new BoolData(false),
                        false,
                        "If the document has revisions that are not accepted, consider them as accepted for the conversion",
                        false
                    )
                },
                { "pagination",
                    new ScriptParameter(
                        "pagination",
                        "Pagination mode",
                        PageNumberingChoice.DataType(),
                        false,
                        "Define how page numbers are computed and inserted in the result",
                        false // from settings
                    )
                },
                { "image-size",
                    new ScriptParameter(
                        "image-size",
                        "Image resizing",
                        ImageOptionChoice.DataType(),
                        false,
                        "",
                        false // from settings
                    )
                },
                { "dpi",
                    new ScriptParameter(
                        "dpi",
                        "Image resampling value",
                        new EnumData(new Dictionary<string, object>()
                        {
                            { "96", 96 },
                            { "120", 120 },
                            { "300", 300 }
                        }, Instance.ImageResamplingValue),
                        false,
                        "Target resolution in dot-per-inch for image resampling, if resampling is requested",
                        false // from settings
                    )
                },
                {
                    "character-styles",
                    new ScriptParameter(
                        "character-styles",
                        "Translate character styles",
                        new BoolData(false),
                        false,
                        "",
                        false // from settings
                    )
                },
                {
                    "footnotes-position",
                    new ScriptParameter(
                        "footnotes-position",
                        "Footnotes position",
                        FootnotesPositionChoice.DataType(),
                        false,
                        "Footnotes position in content",
                        false // from settings
                    )
                },
                {
                    "footnotes-level",
                    new ScriptParameter(
                        "footnotes-level",
                        "Footnotes insertion level",
                        new EnumData(new Dictionary<string, object>()
                        {
                            { "0", 0 },
                            { "1", 1 },
                            { "2", 2 },
                            { "3", 3 },
                            { "4", 4 },
                            { "5", 5 },
                            { "6", 6 },
                        }, Instance.FootnotesLevel),
                        false,
                        "Lowest level into which notes are inserted in content. 0 means the footnotes will be inserted as close as possible of its first call.",
                        false // from settings
                    )
                },
                {
                    "footnotes-numbering",
                    new ScriptParameter(
                        "footnotes-numbering",
                        "Footnotes numbering scheme",
                        FootnotesNumberingChoice.DataType(),
                        false,
                        "Customize footnotes numbering",
                        false // from settings
                    )
                },
                {
                    "footnotes-start-value",
                    new ScriptParameter(
                        "footnotes-start-value",
                        "Footnotes starting value",
                        new IntegerData(min:1,int.MaxValue,Instance.FootnotesStartValue),
                        false,
                        "If footnotes numbering is required, start the notes numbering process from this value",
                        false // from settings
                    )
                },
                {
                    "footnotes-numbering-prefix",
                    new ScriptParameter(
                        "footnotes-numbering-prefix",
                        "Footnotes number prefix",
                        new StringData(Instance.FootnotesNumberingPrefix),
                        false,
                        "Add a prefix before the note's number if numbering is requested",
                        false // from settings
                    )
                },
                {
                    "footnotes-numbering-suffix",
                    new ScriptParameter(
                        "footnotes-numbering-suffix",
                        "Footnotes number suffix",
                        new StringData(Instance.FootnotesNumberingSuffix),
                        false,
                        "Add a text between the note's number and the note content.",
                        false // from settings
                    )
                },
                {"tts-config", new ScriptParameter(
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
                },
                // NP 2025/10/13 : word to dtbook now embed the cleaning steps
                {
                  "repair",
                  new ScriptParameter(
                    "repair",
                    "Repair the dtbook",
                    new BoolData(true),
                    true,
                    ""
                  )
                },
                {
                  "tidy",
                  new ScriptParameter(
                    "tidy",
                    "Tidy up the dtbook",
                    new BoolData(true),
                    true,
                    ""

                  )
                },
                {
                  "folder-depth",
                  new ScriptParameter(
                    "folder-depth",
                    "Folder depth",
                    new EnumData(
                        new Dictionary<string,object>(){
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
            return Directory.GetFiles(inputDirectory.FullName, "*.docx", SearchOption.AllDirectories)[0];
        }

        public override void ExecuteScript(string inputPath)
        {
            base.ExecuteScript(inputPath);
            if(ExtractedShapes.Count > 0) {
                foreach (string shape in ExtractedShapes)
                {
                    File.Copy(shape, Path.Combine(Parameters["output"].Value.ToString(), Path.GetFileName(shape)), true);
                }
            }
        }
    }
}
