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
    public class WordToDtbook : Pipeline2Script
    {
        public WordToDtbook(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "word-to-dtbook";
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
                { "title",
                    new ScriptParameter(
                        "title",
                        "Document title",
                        new StringDataType(),
                        "",
                        false
                    )
                },
                { "creator",
                    new ScriptParameter(
                        "creator",
                        "Document creator or author",
                        new StringDataType(),
                        "",
                        false
                    )
                },
                { "publisher",
                    new ScriptParameter(
                        "publisher",
                        "Document publisher",
                        new StringDataType(),
                        "",
                        false
                    )
                },
                { "uid",
                    new ScriptParameter(
                        "uid",
                        "Document identifier",
                        new StringDataType(),
                        "",
                        false,
                        "Identifier to be added as dtb:uid metadata"
                    )
                },
                { "subject",
                    new ScriptParameter(
                        "subject",
                        "Subject(s)",
                        new StringDataType(),
                        "",
                        false,
                        "Subject(s) to be added as dc:Subject metadata"
                    )
                },
                { "accept-revisions",
                    new ScriptParameter(
                        "accept-revisions",
                        "Accept revisions",
                        new BoolDataType(true),
                        true,
                        false,
                        "If the document has revisions that are not accepted, consider them as accepted for the conversion"
                    )
                },
                { "pagination",
                    new ScriptParameter(
                        "pagination",
                        "Pagination mode",
                        PageNumberingChoice.DataType,
                        PageNumberingChoice.Values[Instance.PagenumStyle],
                        false,
                        "Define how page numbers are computed and inserted in the result",
                        false // from settings
                    )
                },
                { "image-size",
                    new ScriptParameter(
                        "image-size",
                        "Image resizing",
                        ImageOptionChoice.DataType,
                        ImageOptionChoice.Values[Instance.ImageOption],
                        false,
                        "",
                        false // from settings
                    )
                },
                { "dpi",
                    new ScriptParameter(
                        "dpi",
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
                    "character-styles",
                    new ScriptParameter(
                        "character-styles",
                        "Translate character styles",
                        new BoolDataType(false),
                        false,
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
                        FootnotesPositionChoice.DataType,
                        FootnotesPositionChoice.Values[Instance.FootnotesPosition],
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
                    "footnotes-numbering",
                    new ScriptParameter(
                        "footnotes-numbering",
                        "Footnotes numbering scheme",
                        FootnotesNumberingChoice.DataType,
                        FootnotesNumberingChoice.Values[Instance.FootnotesNumbering],
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
                        new IntegerDataType(min:1),
                        Instance.FootnotesStartValue,
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
                        new StringDataType(),
                        Instance.FootnotesNumberingPrefix,
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
                        new StringDataType(),
                        Instance.FootnotesNumberingSuffix,
                        false,
                        "Add a text between the note's number and the note content.",
                        false // from settings
                    )
                },
                {
                    "extract-shapes",
                    new ScriptParameter(
                        "extract-shapes",
                        "Extract shapes",
                        new BoolDataType(),
                        false,
                        false,
                        "",
                        false // hidden
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
    }
}
