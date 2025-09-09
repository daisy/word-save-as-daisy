using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class WordToDaisy3 : Script
    {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        List<Script> scripts;


        public WordToDaisy3(IConversionEventsHandler e) : base(e)
        {
            this.niceName = "Export to DAISY3";
            scripts = new List<Script>()
            {
                new WordToDtbook(e),
                new DtbookCleaner(e),
                new DtbookToDaisy3(e)
            };

            // preset parameters for the cleaning script
            scripts[1].Parameters["tidy"].Value = true;
            scripts[1].Parameters["repair"].Value = true;
            scripts[1].Parameters["narrator"].Value = true;

            StepsCount = scripts.Count;

            // use dtbook to daisy3 parameters
            _parameters = new Dictionary<string, ScriptParameter> {
                {"input", new ScriptParameter(
                        "source",
                        "input",
                        new PathData(PathData.InputOrOutput.input,PathData.FileOrDirectory.File),
                        true,
                        "input",
                        false,
                        ParameterDirection.Input
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
                //{"validation-status", new ScriptParameter(
                //    "validation-status",
                //    "validation-status"
                //    )
                //},
                {"include-tts-log", new ScriptParameter(
                        "include-tts-log",
                        "include-tts-log",
                        new BoolData(),
                        false,
                        "Include tts log with the result",
                        true
                    )
                },
                {"tts-log", new ScriptParameter(
                        "tts-log",
                        "TTS log output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        false,
                        "TTS log output directory",
                         false,
                        ParameterDirection.Output
                    )
                },
                {"publisher", new ScriptParameter(
                        "publisher",
                        "Publisher",
                        new StringData(""),
                        false,
                        "The agency responsible for making the Digital Talking Book available. If left blank, it will be retrieved from the DTBook meta-data.",
                        false
                    )
                },
                {"output", new ScriptParameter(
                        "output-dir",
                        "Daisy3 output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "The resulting DAISY 3 publication."
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
                        new BoolData(true),
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
                {"with-text", new ScriptParameter(
                        "with-text",
                        "With text",
                        new BoolData(true),
                        false,
                        "Includes DTBook in output, as opposed to audio only."
                    )
                },
                {"word-detection", new ScriptParameter(
                        "word-detection",
                        "Apply word detection",
                        new BoolData(false),
                        false,
                        "Whether to detect and mark up words with <w> tags.\r\n" +
                        "\r\n" +
                        "By default word detection is performed but an option is provided to disable it because some DAISY 3\r\n" +
                        "reading systems can't handle the word tags."
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
                        "{0}_DAISY3_{1}",
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
                this.EventsHandler.OnConversionError(new Exception("An error occurred while executing the Word to DAISY 3 conversion pipeline.", ex));
                throw ex;
            }
        }

        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}
