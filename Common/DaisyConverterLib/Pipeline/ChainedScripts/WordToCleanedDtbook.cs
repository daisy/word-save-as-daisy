using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts
{
    public class WordToCleanedDtbook : Script
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        Script wordToDtbook;
        Script dtbookCleaner;

        public WordToCleanedDtbook(IConversionEventsHandler e) : base(e)
        {
            this.niceName = "Export to DTBook XML";
            wordToDtbook = new WordToDtbook(e);
            dtbookCleaner = new DtbookCleaner(e);
            // TODO : for now we consider the 3 global steps of the progression but some granularity within
            // scripts could be taked in account
            StepsCount = 4;
            // set dtbook cleaner to apply default cleanups
            dtbookCleaner.Parameters["tidy"].ParameterValue = true;
            dtbookCleaner.Parameters["repair"].ParameterValue = true;
            dtbookCleaner.Parameters["narrator"].ParameterValue = false;
            // use dtbook to epub3 parameters
            _parameters = new Dictionary<string, ScriptParameter>
            {
                {
                    "input", new ScriptParameter(
                        "source",
                        "Input Docx file",
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File
                        ),
                        "",
                        true,
                        "The document you want to convert."
                    )
                },
                {
                    "output", new ScriptParameter(
                        "result",
                        "DTBook output",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "Output folder of the conversion to DTBook XML"
                    )
                },
                { "Title",
                    new ScriptParameter(
                        "Title",
                        "Document title",
                        new StringDataType(),
                        "",
                        false,"",false
                    )
                },
                { "Creator",
                    new ScriptParameter(
                        "Creator",
                        "Document creator or author",
                        new StringDataType(),
                        "",
                        false,"",false
                    )
                },
                { "Publisher",
                    new ScriptParameter(
                        "Publisher",
                        "Document publisher",
                        new StringDataType(),
                        "",
                        false,"",false
                    )
                },
                { "UID",
                    new ScriptParameter(
                        "UID",
                        "Document identifier",
                        new StringDataType(),
                        "",
                        false,
                        "Identifier to be added as dtb:uid metadata",
                        false
                    )
                },
                { "Subject",
                    new ScriptParameter(
                        "Subject",
                        "Subject(s)",
                        new StringDataType(),
                        "",
                        false,
                        "Subject(s) to be added as dc:Subject metadata",
                        false
                    )
                },
                { "acceptRevisions",
                    new ScriptParameter(
                        "acceptRevisions",
                        "Accept revisions",
                        new BoolDataType(true),
                        true,
                        false,
                        "If the document has revisions that are not accepted, consider them as accepted for the conversion",
                        false
                    )
                },
                { "PagenumStyle",
                    new ScriptParameter(
                        "PagenumStyle",
                        "Pagination mode",
                        PageNumberingChoice.DataType,
                        PageNumberingChoice.Values[Instance.PagenumStyle],
                        false,
                        "Define how page numbers are computed and inserted in the result",
                        false // from settings
                    )
                },
                { "ImageSizeOption",
                    new ScriptParameter(
                        "ImageSizeOption",
                        "Image resizing",
                        ImageOptionChoice.DataType,
                        ImageOptionChoice.Values[Instance.ImageOption],
                        false,
                        "",
                        false // from settings
                    )
                },
                { "DPI",
                    new ScriptParameter(
                        "DPI",
                        "Image resampling value",
                        new EnumDataType(new Dictionary<string, object>()
                        {
                            { "96", 96 },
                            { "120", 120 },
                            { "300", 300 }
                        }, "96"),
                        Instance.ImageResamplingValue,
                        false,
                        "Target resolution in dot-per-inch for image resampling, if resampling is requested",
                        false // from settings
                    )
                },
                {
                    "CharacterStyles",
                    new ScriptParameter(
                        "CharacterStyles",
                        "Translate character styles",
                        new BoolDataType(false),
                        false,
                        false,
                        "",
                        false // from settings
                    )
                },
                {
                    "FootnotesPosition",
                    new ScriptParameter(
                        "FootnotesPosition",
                        "Footnotes position",
                        FootnotesPositionChoice.DataType,
                        FootnotesPositionChoice.Values[Instance.FootnotesPosition],
                        false,
                        "Footnotes position in content",
                        false // from settings
                    )
                },
                {
                    "FootnotesLevel",
                    new ScriptParameter(
                        "FootnotesLevel",
                        "Footnotes insertion level",
                        new EnumDataType(new Dictionary<string, object>()
                        {
                            { "0", 0 },
                            { "1", 1 },
                            { "2", 2 },
                            { "3", 3 },
                            { "4", 4 },
                            { "5", 5 },
                            { "6", 6 },
                        }, "0"),
                        Instance.FootnotesLevel,
                        false,
                        "Lowest level into which notes are inserted in content. 0 means the footnotes will be inserted as close as possible of its first call.",
                        false // from settings
                    )
                },
                {
                    "FootnotesNumbering",
                    new ScriptParameter(
                        "FootnotesNumbering",
                        "Footnotes numbering scheme",
                        FootnotesNumberingChoice.DataType,
                        FootnotesNumberingChoice.Values[Instance.FootnotesNumbering],
                        false,
                        "Customize footnotes numbering",
                        false // from settings
                    )
                },
                {
                    "FootnotesStartValue",
                    new ScriptParameter(
                        "FootnotesStartValue",
                        "Footnotes starting value",
                        new IntegerDataType(min:1),
                        Instance.FootnotesStartValue,
                        false,
                        "If footnotes numbering is required, start the notes numbering process from this value",
                        false // from settings
                    )
                },
                {
                    "FootnotesNumberingPrefix",
                    new ScriptParameter(
                        "FootnotesNumberingPrefix",
                        "Footnotes number prefix",
                        new StringDataType(),
                        Instance.FootnotesNumberingPrefix,
                        false,
                        "Add a prefix before the note's number if numbering is requested",
                        false // from settings
                    )
                },
                {
                    "FootnotesNumberingSuffix",
                    new ScriptParameter(
                        "FootnotesNumberingSuffix",
                        "Footnotes number suffix",
                        new StringDataType(),
                        Instance.FootnotesNumberingSuffix,
                        false,
                        "Add a text between the note's number and the note content.",
                        false // from settings
                    )
                },
                {
                    "extractShapes",
                    new ScriptParameter(
                        "extractShapes",
                        "extractShapes",
                        new BoolDataType(),
                        false,
                        false,
                        "",
                        false // hidden
                    )
                },
            };
        }

        public override void ExecuteScript(string inputPath, bool isQuite)
        {

            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                string.Format(
                    "{0}_DTBookXML_{1}",
                    Path.GetFileNameWithoutExtension(inputPath),
                    DateTime.Now.ToString("yyyyMMddHHmmssffff")
                )
            );
            // Remove and recreate result folder
            // Since the DaisyToEpub3 requires output folder to be empty
            if (Directory.Exists(finalOutput)) {
                Directory.Delete(finalOutput, true);
            }
            Directory.CreateDirectory(finalOutput);
            // transfer parameters value
            foreach (var kv in this._parameters) {
                if (wordToDtbook.Parameters.ContainsKey(kv.Key)) {
                    wordToDtbook.Parameters[kv.Key] = kv.Value;
                }
            }

            DirectoryInfo tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()));
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning " + this._parameters["input"].ParameterValue + " into " + tempDir.FullName));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning the DTBook XML... "));
#endif
            wordToDtbook.Parameters["input"].ParameterValue = this._parameters["input"].ParameterValue;
            wordToDtbook.Parameters["output"].ParameterValue = tempDir.FullName;
            wordToDtbook.ExecuteScript(inputPath, true);

            

            // rebind input and output
            try {
                dtbookCleaner.Parameters["input"].ParameterValue = Directory.GetFiles(tempDir.FullName, "*.xml", SearchOption.AllDirectories)[0];
            }
            catch {
                throw new FileNotFoundException("Could not find result of cleaning process in result folder", tempDir.FullName);
            }
            dtbookCleaner.Parameters["output"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "DTBook XML")).FullName;
            
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning " + dtbookCleaner.Parameters["input"].ParameterValue + " dtbook XML into " + dtbookCleaner.Parameters["output"].ParameterValue));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting DTBook XML to EPUB3..."));
#endif

            dtbookCleaner.ExecuteScript(dtbookCleaner.Parameters["input"].ParameterValue.ToString());

        }
    }
}
