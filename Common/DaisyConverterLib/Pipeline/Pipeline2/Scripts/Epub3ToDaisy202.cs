using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class Epub3ToDaisy202 : Pipeline2Script
    {
        public Epub3ToDaisy202(IConversionEventsHandler eventsHandler) : base(eventsHandler)
        {
            this.name = "epub3-to-daisy202";
            _parameters = new Dictionary<string, ScriptParameter>
            {
                { "input", new ScriptParameter(
                        "source",
                        "EPUB 3 Publication",
                        new PathData(
                            PathData.InputOrOutput.input,
                            PathData.FileOrDirectory.File,
                            "application/epub+zip"
                        ),
                        true,
                        "The package file of the input DTB.",
                        true,
                        ParameterDirection.Input
                    )
                },
                {"output", new ScriptParameter(
                        "result",
                        "EPUB output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "The produced EPUB."
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
            return Directory.GetFiles(inputDirectory.FullName, "*.epub", SearchOption.AllDirectories)[0];
        }
    }
}
