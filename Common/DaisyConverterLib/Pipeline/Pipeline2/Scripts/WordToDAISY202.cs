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
    public class WordToDAISY202 : Pipeline2Script
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public WordToDAISY202(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-daisy202";
            this.niceName = "Export to DAISY2.02";
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
                        "DAISY2.02 output",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "Output folder of the conversion to DAISY2.02"
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
                //{
                //    "extract-shapes",
                //    new ScriptParameter(
                //        "extract-shapes",
                //        "Extract shapes",
                //        new BoolData(false),
                //        false,
                //        "",
                //        false // hidden
                //    )
                //},
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
                { "validation", new ScriptParameter(
                        "validation",
                        "Validation",
                        new EnumData(
                            new Dictionary<string, object> {
                                { "No validation", "off" },
                                { "Report validation issues", "report" },
                                { "Abort on validation issues", "abort" },
                            }, "abort"),
                        false,
                        "Whether to abort on validation issues."
                    )
                },
                { "validation-report", new ScriptParameter(
                        "validation-report",
                        "Validation reports",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        false,
                        "Output path of the validation reports",
                        false,
                        ParameterDirection.Output
                    )
                },
                {"language", new ScriptParameter(
                        "language",
                        "Language code",
                        new StringData(),
                        false,
                        "Language code of the input document."
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
                {"audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolData(false),
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
                //{
                //  "narrator",
                //  new ScriptParameter(
                //    "narrator",
                //    "Prepare dtbook for pipeline 1 narrator",
                //    new BoolData(false),
                //    true,
                //    ""

                //  )
                //},
                //{
                //  "ApplySentenceDetection",
                //  new ScriptParameter(
                //    "ApplySentenceDetection",
                //    "Apply sentences detection",
                //    new BoolData(false),
                //    true,
                //    ""
                //  )
                //},
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
