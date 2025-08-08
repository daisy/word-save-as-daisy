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
        List<Script> scripts;

        public WordToCleanedDtbook(IConversionEventsHandler e) : base(e)
        {
            this.niceName = "Export to DTBook XML";
            scripts = new List<Script>() {
                new WordToDtbook(e),
                new DtbookCleaner(e)
            };
            StepsCount = scripts.Count;

            // use dtbook to epub3 parameters
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
                        "DTBook output",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "Output folder of the conversion to DTBook XML"
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
                  "narrator",
                  new ScriptParameter(
                    "narrator",
                    "Prepare dtbook for pipeline 1 narrator",
                    new BoolData(false),
                    true,
                    ""

                  )
                },
            };
        }

        public override void ExecuteScript(string inputPath)
        {
            try {
                // Create a directory using the document name
                DirectoryInfo finalOutput = new DirectoryInfo(
                    Path.Combine(
                    Parameters["output"].Value.ToString(),
                    string.Format(
                        "{0}_DTBookXML_{1}",
                        Path.GetFileNameWithoutExtension(inputPath),
                        DateTime.Now.ToString("yyyyMMddHHmmssffff")
                    )
                ));
                // Remove and recreate result folder
                // Since the DaisyToEpub3 requires output folder to be empty
                if (finalOutput.Exists) {
                    finalOutput.Delete(true);
                    finalOutput.Create();
                }

                string input = inputPath;
                DirectoryInfo outputDir = finalOutput;

                for (int i = 0; i < scripts.Count; i++) {
                    if (i > 0) {
                        // chain last output to next input for non-first scripts
                        try {
                            input = scripts[i].searchInputFromDirectory(outputDir);
                        }
                        catch {
                            throw new FileNotFoundException($"Could not find result of previous script {scripts[i - 1].Name} in intermediate folder", outputDir.FullName);
                        }
                    }
                    // create a temporary output directory for all scripts except the last one
                    outputDir = i < scripts.Count - 1
                             ? Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()))
                             : finalOutput;
                    // transfer global parameters value except input and output (that could change between scripts)
                    foreach (var k in this._parameters.Keys.Except(new string[] { "input", "output" })) {
                        if (scripts[i].Parameters.ContainsKey(k)) {
                            scripts[i].Parameters[k] = this._parameters[k];
                        }
                    }

#if DEBUG
                    this.EventsHandler.onProgressMessageReceived(
                        this,
                        new DaisyEventArgs(
                            $"Applying {scripts[i].Name} on {input} and storing into {outputDir.FullName}"
                        )
                    );
#else
                this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs($"Launching script {scripts[i].Name} ... "));
#endif
                    // rebind input and output
                    scripts[i].Parameters["input"].Value = input;
                    scripts[i].Parameters["output"].Value = outputDir.FullName;
                    scripts[i].ExecuteScript(inputPath);
                }
                System.Diagnostics.Process.Start(finalOutput.FullName);
            }
            catch (Exception ex) {
                this.EventsHandler.OnConversionError(new Exception("An error occurred while executing the Word to DTBook XML conversion pipeline.", ex));
                throw ex;
            }
            
        }

        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}
