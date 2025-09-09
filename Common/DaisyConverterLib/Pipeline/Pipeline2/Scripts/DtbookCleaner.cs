using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts
{
    public class DtbookCleaner : Pipeline2Script
    {
        public DtbookCleaner(IConversionEventsHandler eventsHandler) : base(eventsHandler)
        {
            this.name = "dtbook-cleaner";
            this.niceName = "Export to dtbook XML";
            _parameters = new Dictionary<string, ScriptParameter>
            {
                {
                    "input",
                    new ScriptParameter(
                        "source",
                        "DTBook file(s)",
                        new PathData(
                            PathData.InputOrOutput.input,
                            PathData.FileOrDirectory.File,
                            "application/x-dtbook+xml",
                            ""
                        ),
                        true,
                        "One or more DTBook files to be cleaned"
                    )
                },
                {
                    "output",
                    new ScriptParameter(
                        "result",
                        "DTBook XML output folder",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory,
                            "application/x-dtbook+xml",
                            ""
                        ),
                        true,
                        "Cleaned DTBooks"
                    )
                },
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
                  "simplifyHeadingLayout",
                  new ScriptParameter(
                    "simplifyHeadingLayout",
                    "Tidy - Simplify headings layout",
                    new BoolData(false),
                    false,
                    "",
                    false

                  )
                },
                {
                  "externalizeWhitespace",
                  new ScriptParameter(
                    "externalizeWhitespace",
                    "Tidy - Externalize whitespaces",
                    new BoolData(false),
                    false,
                    "",
                    false

                  )
                },
                {
                  "documentLanguage",
                  new ScriptParameter(
                    "documentLanguage",
                    "Tidy - Document language",
                    new StringData(""),
                    false,
                    "",
                    false
                  )
                },
                {
                  "narrator",
                  new ScriptParameter(
                    "narrator",
                    "Prepare dtbook for pipeline 1 narrator",
                    new BoolData(false),
                    true,
                    ""

                  )
                },
                {
                  "publisher",
                  new ScriptParameter(
                    "publisher",
                    "Narrator - Publisher",
                    new StringData(""),
                    false,
                    "",
                    false

                  )
                },
                {
                  "ApplySentenceDetection",
                  new ScriptParameter(
                    "ApplySentenceDetection",
                    "Apply sentences detection",
                    new BoolData(false),
                    true,
                    ""
                  )
                },
                {
                  "WithDoctype",
                  new ScriptParameter(
                    "WithDoctype",
                    "Include the doctype in resulting dtbook(s)",
                    new BoolData(true),
                    false,
                    "",
                    false
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