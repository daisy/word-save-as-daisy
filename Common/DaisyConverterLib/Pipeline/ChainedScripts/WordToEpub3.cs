using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class WordToEpub3 : Script {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;


        List<Pipeline2Script> scripts;


        public WordToEpub3(IConversionEventsHandler e): base(e) {
            this.niceName = "Export to EPUB3";

            scripts = new List<Pipeline2Script>() {
                new WordToDtbook(e),
                new DtbookCleaner(e),
                new DtbookToEpub3(e)
            };
            StepsCount = scripts.Count;

            // preset parameters for the cleaning script
            scripts[1].Parameters["tidy"].ParameterValue = true;
            scripts[1].Parameters["repair"].ParameterValue = true;
            scripts[1].Parameters["narrator"].ParameterValue = true;

            // use dtbook to epub3 parameters
            _parameters = new Dictionary<string, ScriptParameter>
            {
                { "input", new ScriptParameter(
                        "source",
                        "OPF",
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File,
                            "application/oebps-package+xml"
                        ),
                        "",
                        true,
                        "The package file of the input DTB.",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                {"include-tts-log", new ScriptParameter(
                        "include-tts-log",
                        "include-tts-log",
                        new BoolDataType(),
                        false,
                        false,
                        "Include tts log with the result",
                        true
                    )
                },
                {"language", new ScriptParameter(
                        "language",
                        "Language code",
                        new StringDataType(),
                        "",
                        false,
                        "Language code of the input document."
                    )
                },
                {"output", new ScriptParameter(
                        "result",
                        "EPUB output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The produced EPUB."
                    )
                },
                { "title",
                    new ScriptParameter(
                        "title",
                        "Document title",
                        new StringDataType(),
                        "",
                        false,"",false
                    )
                },
                { "creator",
                    new ScriptParameter(
                        "creator",
                        "Document creator or author",
                        new StringDataType(),
                        "",
                        false,"",false
                    )
                },
                { "uid",
                    new ScriptParameter(
                        "uid",
                        "Document identifier",
                        new StringDataType(),
                        "",
                        false,
                        "Identifier to be added as dtb:uid metadata",
                        false
                    )
                },
                { "subject",
                    new ScriptParameter(
                        "subject",
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
                { "validation", new ScriptParameter(
                        "validation",
                        "Validation",
                        new EnumDataType(
                            new Dictionary<string, object> {
                                { "No validation", "off" },
                                { "Report validation issues", "report" },
                                { "Abort on validation issues", "abort" },
                            }, "Abort on validation issues"),
                        "abort",
                        false,
                        "Whether to abort on validation issues."
                    )
                },
                { "validation-report", new ScriptParameter(
                        "validation-report",
                        "Validation reports",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        false,
                        "Output path of the validation reports",
                        false,
                        ScriptParameter.ParameterDirection.Output
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
                        new BoolDataType(),
                        false,
                        false,
                        "Whether the input DTBook is a NIMAS 1.1-conformant XML content file.",
                        false // not sure this option should available in save as daisy
                    )
                },
                {"tts-config", new ScriptParameter(
                        "tts-config",
                        "Text-to-speech configuration file",
                        new PathDataType(PathDataType.InputOrOutput.input,PathDataType.FileOrDirectory.File, "", GlobaleSettings.TTSConfigFile ?? ""),
                        GlobaleSettings.TTSConfigFile ?? "",
                        false,
                        "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                        true,
                        ScriptParameter.ParameterDirection.Input
                    )
                },
                {"audio", new ScriptParameter(
                        "audio",
                        "Enable text-to-speech",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
            };
        }

        public override void ExecuteScript(string inputPath, bool isQuite) {
            // Create a directory using the document name
            DirectoryInfo finalOutput = new DirectoryInfo(
                Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                string.Format(
                    "{0}_EPUB3_{1}",
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
                if( scripts[i].Parameters.ContainsKey("validation") &&
                    scripts[i].Parameters.ContainsKey("validation-report") &&
                    scripts[i].Parameters["validation"].ParameterValue.ToString() == "report"
                ) {
                    scripts[i].Parameters["validation-report"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput.FullName, "report")).FullName;
                }
                if (scripts[i].Parameters.ContainsKey("include-tts-log") &&
                    scripts[i].Parameters.ContainsKey("tts-log") &&
                    (bool)scripts[i].Parameters["include-tts-log"].ParameterValue == true
                ) {
                    scripts[i].Parameters["tts-log"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput.FullName, "tts-log")).FullName;
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
                scripts[i].Parameters["input"].ParameterValue = input;
                scripts[i].Parameters["output"].ParameterValue = outputDir.FullName;
                scripts[i].ExecuteScript(inputPath, isQuite);
            }

        }
    }
}
