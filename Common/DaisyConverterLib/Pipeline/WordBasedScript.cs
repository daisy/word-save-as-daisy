using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public class WordBasedScript : Script
    {
        public WordBasedScript(IConversionEventsHandler e)
            : base(e)
        {
            _alsoExportShapes = true;
            _parameters.Add("input",
                new ScriptParameter(
                    "source",
                    "Input Docx file",
                    new PathData(
                        PathData.InputOrOutput.input,
                        PathData.FileOrDirectory.File
                    ),
                    true,
                    "The document you want to convert."
                )
            );
            _parameters.Add("title",
                new ScriptParameter(
                    "title",
                    "Document title",
                    new StringData(),
                    false, "", false
                )
            );
            _parameters.Add("creator",
                new ScriptParameter(
                    "creator",
                    "Document creator or author",
                    new StringData(),
                    false, "", false
                )
            );
            _parameters.Add("publisher",
                new ScriptParameter(
                    "publisher",
                    "Document publisher",
                    new StringData(),
                    false, "", false
                )
            );
            _parameters.Add("uid",
                new ScriptParameter(
                    "uid",
                    "Document identifier",
                    new StringData(),
                    false,
                    "Identifier to be added as dtb:uid metadata",
                    false
                )
            );
            _parameters.Add("subject",
                new ScriptParameter(
                    "subject",
                    "Subject(s)",
                    new StringData(),
                    false,
                    "Subject(s) to be added as dc:Subject metadata",
                    false
                )
            );
            _parameters.Add("accept-revisions",
                new ScriptParameter(
                    "accept-revisions",
                    "Accept revisions",
                    new BoolData(false),
                    false,
                    "If the document has revisions that are not accepted, consider them as accepted for the conversion",
                    false
                )
            );
            _parameters.Add("pagination",
                new ScriptParameter(
                    "pagination",
                    "Pagination mode",
                    PageNumberingChoice.DataType(),
                    false,
                    "Define how page numbers are computed and inserted in the result",
                    false // from settings
                )
            );
            _parameters.Add("image-size",
                new ScriptParameter(
                    "image-size",
                    "Image resizing",
                    ImageOptionChoice.DataType(),
                    false,
                    "",
                    false // from settings
                )
            );
            _parameters.Add("dpi",
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
            );
            _parameters.Add("character-styles",
                new ScriptParameter(
                    "character-styles",
                    "Translate character styles",
                    new BoolData(false),
                    false,
                    "",
                    false // from settings
                )
            );
            _parameters.Add("footnotes-position",
                new ScriptParameter(
                    "footnotes-position",
                    "Footnotes position",
                    FootnotesPositionChoice.DataType(),
                    false,
                    "Footnotes position in content",
                    false // from settings
                )
            );
            _parameters.Add("footnotes-level",
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
            );
            _parameters.Add("footnotes-numbering",
                new ScriptParameter(
                    "footnotes-numbering",
                    "Footnotes numbering scheme",
                    FootnotesNumberingChoice.DataType(),
                    false,
                    "Customize footnotes numbering",
                    false // from settings
                )
            );
            _parameters.Add("footnotes-start-value",
                new ScriptParameter(
                    "footnotes-start-value",
                    "Footnotes starting value",
                    new IntegerData(min: 1, int.MaxValue, Instance.FootnotesStartValue),
                    false,
                    "If footnotes numbering is required, start the notes numbering process from this value",
                    false // from settings
                )
            );
            _parameters.Add("footnotes-numbering-prefix",
                new ScriptParameter(
                    "footnotes-numbering-prefix",
                    "Footnotes number prefix",
                    new StringData(Instance.FootnotesNumberingPrefix),
                    false,
                    "Add a prefix before the note's number if numbering is requested",
                    false // from settings
                )
            );
            _parameters.Add("footnotes-numbering-suffix",
                new ScriptParameter(
                    "footnotes-numbering-suffix",
                    "Footnotes number suffix",
                    new StringData(Instance.FootnotesNumberingSuffix),
                    false,
                    "Add a text between the note's number and the note content.",
                    false // from settings
                )
            );
            _parameters.Add("extract-shapes",
                new ScriptParameter(
                    "extract-shapes",
                    "Extract shapes",
                    new BoolData(false),
                    false,
                    "",
                    false // hidden
                )
            );
            _parameters.Add("language",
                  new ScriptParameter(
                    "language",
                    "Language code",
                    new EnumData(new Dictionary<string, object>(), isEditable: true),
                    true,
                    "Language code of the input document."
                  )
               );
            // NP 2025/10/13 : word to dtbook now embed the cleaning steps
            _parameters.Add(
              "repair",
              new ScriptParameter(
                "repair",
                "Repair the dtbook",
                new BoolData(true),
                true,
                ""
              )
            );
            _parameters.Add(
              "tidy",
              new ScriptParameter(
                "tidy",
                "Tidy up the dtbook",
                new BoolData(true),
                true,
                ""

              )
            );
            // Reserved narrator for dtbook export
            //_parameters.Add(
            //  "narrator",
            //  new ScriptParameter(
            //    "narrator",
            //    "Prepare dtbook for pipeline 1 narrator",
            //    new BoolData(false),
            //    true,
            //    ""
            //  )
            //);
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
