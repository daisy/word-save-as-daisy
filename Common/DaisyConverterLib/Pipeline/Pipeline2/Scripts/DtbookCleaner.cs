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
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File,
                            "application/x-dtbook+xml"
                        ),
                        "",
                        true,
                        "One or more DTBook files to be cleaned"
                    )
                },
                {
                    "output",
                    new ScriptParameter(
                        "result",
                        "DTBook XML output folder",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory,
                            "application/x-dtbook+xml"
                        ),
                        "",
                        true,
                        "Cleaned DTBooks"
                    )
                },
                {
                  "repair",
                  new ScriptParameter(
                    "repair",
                    "Repair the dtbook",
                    new BoolDataType(true),
                    true,
                    true,
                    ""
                  )
                },
                // Hidden for now, as not implemented in pipeline 2 script
                //{
                //  "fixCharset",
                //  new ScriptParameter(
                //    "fixCharset",
                //    "Repair - Fix Charset",
                //    new BoolDataType(false),
                //    false,
                //    false,
                //    "",
                //    false

                //  )
                //},
                {
                  "tidy",
                  new ScriptParameter(
                    "tidy",
                    "Tidy up the dtbook",
                    new BoolDataType(false),
                    true,
                    true,
                    ""

                  )
                },
                {
                  "simplifyHeadingLayout",
                  new ScriptParameter(
                    "simplifyHeadingLayout",
                    "Tidy - Simplify headings layout",
                    new BoolDataType(false),
                    false,
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
                    new BoolDataType(false),
                    false,
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
                    new StringDataType(""),
                    "",
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
                    new BoolDataType(false),
                    false,
                    true,
                    ""

                  )
                },
                {
                  "publisher",
                  new ScriptParameter(
                    "publisher",
                    "Narrator - Publisher",
                    new StringDataType(""),
                    "",
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
                    new BoolDataType(false),
                    false,
                    true,
                    ""
                  )
                },
                {
                  "WithDoctype",
                  new ScriptParameter(
                    "WithDoctype",
                    "Include the doctype in resulting dtbook(s)",
                    new BoolDataType(false),
                    true,
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
            return Directory.GetFiles(inputDirectory.FullName, "*.xml", SearchOption.AllDirectories)[0];
        }
    }
}