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
                {
                  "fixCharset",
                  new ScriptParameter(
                    "fixCharset",
                    "Repair - Fix Charset",
                    new BoolDataType(false),
                    false,
                    false,
                    "",
                    false

                  )
                },
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

        public override void ExecuteScript(string inputPath, bool isQuite)
        {
            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                string.Format(
                    "{0}_DtbookXML_{1}",
                    Path.GetFileNameWithoutExtension(inputPath),
                    DateTime.Now.ToString("yyyyMMddHHmmssffff")
                )
            );
            // Remove and recreate result folder
            if (Directory.Exists(finalOutput))
            {
                Directory.Delete(finalOutput, true);
            }
            Directory.CreateDirectory(finalOutput);
            // Update final output with the new subdirectory
            Parameters["output"].ParameterValue = finalOutput;
            // Execute the script
            base.ExecuteScript(inputPath, isQuite);
            // TODO : fix the resource copy in pipeline 2 script later :
            string scriptInput = this._parameters["input"].ParameterValue.ToString();
            string currentDirectory = Path.GetDirectoryName(scriptInput);
            string currentFile = Path.GetFileName(scriptInput);
            string outputDir = this._parameters["output"].ParameterValue.ToString();

            string[] files = Directory.GetFiles(currentDirectory, "*", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                if(Path.GetFileName(file) != currentFile)
                {
                    string outputFile = Path.Combine(
                        outputDir,
                        Path.GetFileName(file)
                    );
                    File.Copy(file,outputFile,true);
                }
            }
            
        }
    }
}