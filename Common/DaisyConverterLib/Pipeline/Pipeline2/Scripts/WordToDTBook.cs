using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class WordToDTBook : Pipeline2Step
    {
        public WordToDTBook()
        {
            this.nameOrURI = "http://www.daisy.org/pipeline/modules/word-to-dtbook/oox2Daisy.xpl";
            this.niceName = "Word to DTBook XML";
            _parameters = new Dictionary<string, ScriptParameter>();
            _parameters.Add(
                "input",
                new ScriptParameter(
                    "source",
                    "docx",
                    new PathDataType(
                        PathDataType.InputOrOutput.input,
                        PathDataType.FileOrDirectory.File
                    ),
                    "",
                    true,
                    "Input word file"
                )
            );
            _parameters.Add(
                "output",
                new ScriptParameter(
                    "document-output-dir",
                    "XML output directory",
                    new PathDataType(
                        PathDataType.InputOrOutput.output,
                        PathDataType.FileOrDirectory.Directory
                    ),
                    "",
                    true,
                    "Ouput directory for the xml"
                )
            );
            _parameters.Add(
                "output-file",
                new ScriptParameter(
                    "document-output-file",
                    "XML output file path",
                    new PathDataType(
                        PathDataType.InputOrOutput.output,
                        PathDataType.FileOrDirectory.File
                    ),
                    "",
                    true,
                    "Ouput file for the xml"
                )
            );
            //_parameters.Add(
            //    "tempSource",
            //    new ScriptParameter(
            //        "tempSource",
            //        "Temporary working document copy",
            //        new PathDataType(
            //            PathDataType.InputOrOutput.input,
            //            PathDataType.FileOrDirectory.File
            //        ),
            //        "",
            //        false,
            //        "Optional temp file"
            //    )
            //);
            _parameters.Add(
                "pipeline-output",
                new ScriptParameter(
                    "pipeline-output-dir",
                    "pipeline output directory",
                    new PathDataType(
                        PathDataType.InputOrOutput.output,
                        PathDataType.FileOrDirectory.Directory
                    ),
                    "",
                    false,
                    "The produced EPUB."
                )
            );
            _parameters.Add(
                "Title",
                new ScriptParameter(
                    "Title",
                    "Title",
                     new StringDataType(),
                    "",
                    false,
                    "Title"
                )
            );
            _parameters.Add(
                           "Creator",
                           new ScriptParameter(
                               "Creator",
                               "Creator",
                                new StringDataType(),
                               "",
                               false,
                               "Creator"
                           )
                       );
            _parameters.Add(
                           "Publisher",
                           new ScriptParameter(
                               "Publisher",
                               "Publisher",
                                new StringDataType(),
                               "",
                               false,
                               "Publisher"
                           )
                       );
            _parameters.Add(
                           "UID",
                           new ScriptParameter(
                               "UID",
                               "UID",
                                new StringDataType(),
                               "",
                               false,
                               "UID"
                           )
                       );
            _parameters.Add(
                           "Subject",
                           new ScriptParameter(
                               "Subject",
                               "Subject",
                                new StringDataType(),
                               "",
                               false,
                               "Subject"
                           )
                       );
            _parameters.Add(
                           "prmTRACK",
                           new ScriptParameter(
                               "prmTRACK",
                               "prmTRACK",
                                new StringDataType(),
                               "NoTrack",
                               false,
                               "prmTRACK"
                           )
                       );
            _parameters.Add(
                           "Version",
                           new ScriptParameter(
                               "Version",
                               "Version",
                                new StringDataType(),
                               "14",
                               false,
                               "Version"
                           )
                       );
            _parameters.Add(
                           "Custom",
                           new ScriptParameter(
                               "Custom",
                               "Custom",
                                new StringDataType(),
                               "",
                               false,
                               "Custom"
                           )
                       );
            _parameters.Add(
                           "MasterSub",
                           new ScriptParameter(
                               "MasterSub",
                               "MasterSub",
                                new BoolDataType(false),
                               false,
                               false,
                               "MasterSub"
                           )
                       );
            _parameters.Add(
                           "ImageSizeOption",
                           new ScriptParameter(
                               "ImageSizeOption",
                               "ImageSizeOption",
                                new StringDataType("ImageSizeOption"),
                               "original",
                               false,
                               "ImageSizeOption"
                           )
                       );
            _parameters.Add(
                           "DPI",
                           new ScriptParameter(
                               "DPI",
                               "DPI",
                                new IntegerDataType(),
                               96,
                               false,
                               "ImageSizeOption"
                           )
                       );
            _parameters.Add(
                           "CharacterStyles",
                           new ScriptParameter(
                               "CharacterStyles",
                               "CharacterStyles",
                                new BoolDataType(),
                               false,
                               false,
                               "CharacterStyles"
                           )
                       );
            // FIXME : requires an update of the parser on the daisy calabash-adapter side
            //_parameters.Add(
            //               "MathML",
            //               new ScriptParameter(
            //                   "MathML",
            //                   "MathML",
            //                    new MapDataType(),
            //                    new Dictionary<string, List<string>>()
            //                    {
            //                        {"wdTextFrameStory", new List<string>() },
            //                        {"wdFootnotesStory", new List<string>() },
            //                        {"wdMainTextStory", new List<string>() }
            //                    },
            //                   false,
            //                   "MathML"
            //               )
            //           );
            _parameters.Add(
                           "FootnotesPosition",
                           new ScriptParameter(
                               "FootnotesPosition",
                               "FootnotesPosition",
                                new StringDataType(),
                               "end",
                               false,
                               "FootnotesPosition"
                           )
                       );
            _parameters.Add(
                           "FootnotesLevel",
                           new ScriptParameter(
                               "FootnotesLevel",
                               "FootnotesLevel",
                                new IntegerDataType(),
                               1,
                               false,
                               "FootnotesLevel"
                           )
                       );
            // output port that seems required required
            _parameters.Add(
                           "result",
                           new ScriptParameter(
                               "result",
                               "result",
                                new PathDataType(
                                    PathDataType.InputOrOutput.output,
                                    PathDataType.FileOrDirectory.File
                                ),
                               Path.GetTempFileName(),
                               false,
                               "result",
                               false,
                               ScriptParameter.ParameterDirection.Output
                           )
                       );
        }
    }
}
