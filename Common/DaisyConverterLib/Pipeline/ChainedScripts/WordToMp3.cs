using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class WordToMp3 : Script {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        List<Script> scripts;
        public WordToMp3(IConversionEventsHandler e): base(e) {
            this.niceName = "Export to Megavoice MP3 fileset";
            scripts = new List<Script>() {
                new WordToDtbook(e),
                new DtbookCleaner(e),
                new DtbookToDaisy3(e),
                new Daisy3ToMp3(e)
            };
            // preset parameters for the cleaning script
            scripts[1].Parameters["tidy"].Value = true;
            scripts[1].Parameters["repair"].Value = true;
            scripts[1].Parameters["narrator"].Value = true;

            // preset to create an audio-only intermediate daisy3
            scripts[2].Parameters["audio"].Value = true;
            scripts[2].Parameters["with-text"].Value = false;

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
                        new BoolData(false),
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
                        "result",
                        "Output directory",
                        new PathData(
                            PathData.InputOrOutput.output,
                            PathData.FileOrDirectory.Directory
                        ),
                        true,
                        "The resulting MP3 fileset folder."
                    )
                },
                {"tts-config", new ScriptParameter(
                        "tts-config",
                        "Text-to-speech configuration file",
                        new PathData(PathData.InputOrOutput.input,PathData.FileOrDirectory.File, "", GlobaleSettings.TTSConfigFile ?? ""),
                        false,
                        "Configuration file for the text-to-speech.\r\n\r\n[More details on the configuration file format](http://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/).",
                        true,
                        ParameterDirection.Input
                    )
                },
                // From DAISY3 to MP3 script
                {
                  "folder-depth",
                  new ScriptParameter(
                    "folder-depth",
                    "Folder depth",
                    new EnumData(
                        new Dictionary<string,object>(){
                            {"1", "1"},
                            {"2", "2"},
                            {"3", "3"},
                        }, "1"),
                    false,
                    "The number of folder levels in the produced folder structure.\n\n" +
                    "The book is always, if possible, contained in a single top - level folder with MP3 files or\n" +
                    "sub - folders(files for folder depth 1, sub - folders for folder depths greater than 1) that correspond\n" +
                    "with top-level sections of the book.\n\n" +
                    "If there are more top - level sections than the maximum number of files / folders that a top - level\n" +
                    "folder can contain, the book is divided over multiple top-level folders.Similarly, if the number of\n" +
                    "level - two sections within a top-level section exceeds the maximum number of files / folders that a\n" +
                    "level - two folder can contain, the top-level section is divided over multiple level-two folders."

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
                        "{0}_MegaVoiceMP3_{1}",
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
                this.EventsHandler.OnConversionError(new Exception("An error occurred while executing the Word to EPUB 3 conversion pipeline.", ex));
            }
        }

        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}
