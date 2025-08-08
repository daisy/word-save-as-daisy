using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts
{
    public class WordToDaisy202 : Script
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        List<Script> scripts;

        public WordToDaisy202(IConversionEventsHandler e) : base(e)
        {
            this.niceName = "Export to DAISY 2.02";
            scripts = new List<Script>()
            {
                new WordToDtbook(e),
                new DtbookCleaner(e),
                new DtbookToEpub3(e),
                new Epub3ToDaisy202(e)
            };

            // preset parameters for the cleaning script
            scripts[1].Parameters["tidy"].Value = true;
            scripts[1].Parameters["repair"].Value = true;
            scripts[1].Parameters["narrator"].Value = true;

            StepsCount = scripts.Count;

            // use dtbook to epub3 parameters
            _parameters = new Dictionary<string, ScriptParameter>
            {
                { "input", new ScriptParameter(
                        "source",
                        "OPF",
                        new PathData(
                            PathData.InputOrOutput.input,
                            PathData.FileOrDirectory.File,
                            "application/oebps-package+xml"
                        ),
                        true,
                        "The package file of the input DTB.",
                        true,
                        ParameterDirection.Input
                    )
                },
                {"include-tts-log", new ScriptParameter(
                        "include-tts-log",
                        "include-tts-log",
                        new BoolData(false),
                        false,
                        "Include tts log with the result",
                        true
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
                //{ "validation-status", new ScriptParameter(
                //        "validation-report",
                //        "Validation reports",
                //        new PathDataType(
                //            PathDataType.InputOrOutput.output,
                //            PathDataType.FileOrDirectory.Directory
                //        ),
                //        "",
                //        false,
                //        ""
                //    )
                //},
                {"nimas", new ScriptParameter(
                        "nimas",
                        "NIMAS input",
                        new BoolData(),
                        false,
                        "Whether the input DTBook is a NIMAS 1.1-conformant XML content file.",
                        false // not sure this option should available in save as daisy
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
                        true,
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
                        "{0}_DAISY202_{1}",
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
                    if (scripts[i].Parameters.ContainsKey("validation") &&
                        scripts[i].Parameters.ContainsKey("validation-report") &&
                        scripts[i].Parameters["validation"].Value.ToString() == "report"
                    ) {
                        scripts[i].Parameters["validation-report"].Value = Directory.CreateDirectory(Path.Combine(finalOutput.FullName, "report")).FullName;
                    }
                    if (scripts[i].Parameters.ContainsKey("include-tts-log") &&
                        scripts[i].Parameters.ContainsKey("tts-log") &&
                        (bool)scripts[i].Parameters["include-tts-log"].Value == true
                    ) {
                        scripts[i].Parameters["tts-log"].Value = Directory.CreateDirectory(Path.Combine(finalOutput.FullName, "tts-log")).FullName;
                    }

#if DEBUG
                    this.EventsHandler.onProgressMessageReceived(
                        this,
                        new DaisyEventArgs(
                            $"Applying {scripts[i].Name} on {input} and storing into {outputDir.FullName}"
                        )
                    );
#else
                    this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs($"Applying script {scripts[i].Name} ... "));
#endif
                    // rebind input and output
                    scripts[i].Parameters["input"].Value = input;
                    scripts[i].Parameters["output"].Value = outputDir.FullName;
                    scripts[i].ExecuteScript(inputPath);
                }
                System.Diagnostics.Process.Start(finalOutput.FullName);
            }
            catch (Exception ex) {
                this.EventsHandler.OnConversionError(new Exception("An error occurred while executing the Word to DAISY 2.02 conversion pipeline.", ex));
            }
        }

        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}
