using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class CleanedDtbookToMp3 : Script {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        Script dtbookCleaner;
        Script dtbookToDaisy3;
        Script daisy3ToMp3;

        public CleanedDtbookToMp3(IConversionEventsHandler e): base(e) {
            this.niceName = "Export to Megavoice MP3 fileset";
            dtbookCleaner = new DtbookCleaner(e);
            dtbookToDaisy3 = new DtbookToDaisy3(e);
            daisy3ToMp3 = new Daisy3ToMp3(e);
            // TODO : for now we consider the 3 global steps of the progression but some granularity within
            // scripts could be taked in account
            StepsCount = 4;
            // set dtbook cleaner to apply default cleanups
            dtbookCleaner.Parameters["tidy"].ParameterValue = true;
            dtbookCleaner.Parameters["repair"].ParameterValue = true;
            dtbookCleaner.Parameters["narrator"].ParameterValue = true;


            // use dtbook to daisy3 parameters
            _parameters = new Dictionary<string, ScriptParameter> {
                {"input", new ScriptParameter(
                        "source",
                        "input",
                        new PathDataType(PathDataType.InputOrOutput.input,PathDataType.FileOrDirectory.File),
                        "",
                        true,
                        "input",
                        false,
                        ScriptParameter.ParameterDirection.Input
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
                        new BoolDataType(),
                        false,
                        false,
                        "Include tts log with the result",
                        true
                    )
                },
                {"tts-log", new ScriptParameter(
                        "tts-log",
                        "TTS log output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        false,
                        "TTS log output directory",
                         false,
                        ScriptParameter.ParameterDirection.Output
                    )
                },
                {"publisher", new ScriptParameter(
                        "publisher",
                        "Publisher",
                        new StringDataType(""),
                        "",
                        false,
                        "The agency responsible for making the Digital Talking Book available. If left blank, it will be retrieved from the DTBook meta-data.",
                        false
                    )
                },
                {"output", new ScriptParameter(
                        "result",
                        "Output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The resulting MP3 fileset folder."
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
                // From DAISY3 to MP3 script
                {
                  "folder-depth",
                  new ScriptParameter(
                    "folder-depth",
                    "Folder depth",
                    new EnumDataType(
                        new Dictionary<string,object>(){
                            {"1", "1"},
                            {"2", "2"},
                            {"3", "3"},
                        }, "1"),
                    "1",
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

        public override void ExecuteScript(string inputPath, bool isQuite) {

            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                string.Format(
                    "{0}_Megavoice_{1}",
                    Path.GetFileNameWithoutExtension(inputPath),
                    DateTime.Now.ToString("yyyyMMddHHmmssffff")
                )
            );
            // Remove and recreate result folder
            // Since the DaisyToEpub3 requires output folder to be empty
            if (Directory.Exists(finalOutput))
            {
                Directory.Delete(finalOutput, true);
            }
            Directory.CreateDirectory(finalOutput);


            DirectoryInfo tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()));

#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning " + this._parameters["input"].ParameterValue + " into " + tempDir.FullName));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Cleaning the DTBook XML... "));
#endif
            dtbookCleaner.Parameters["input"].ParameterValue = this._parameters["input"].ParameterValue;
            dtbookCleaner.Parameters["output"].ParameterValue = tempDir.FullName;
            dtbookCleaner.ExecuteScript(inputPath, true);


            DirectoryInfo daisy3TempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()));
            // transfer parameters value
            foreach (var kv in this._parameters)
            {
                if (dtbookToDaisy3.Parameters.ContainsKey(kv.Key) && kv.Key != "output") // avoid copy daisy3 to mp3 output as it is not the same named option
                {
                    dtbookToDaisy3.Parameters[kv.Key] = kv.Value;
                }
            }
            dtbookToDaisy3.Parameters["audio"].ParameterValue = true;
            dtbookToDaisy3.Parameters["with-text"].ParameterValue = false;
            // rebind input and output
            try
            {
                dtbookToDaisy3.Parameters["input"].ParameterValue = Directory.GetFiles(tempDir.FullName, "*.xml", SearchOption.AllDirectories)[0];
            }
            catch
            {
                throw new FileNotFoundException("Could not find result of cleaning process in result folder", tempDir.FullName);
            }

            if ((bool)dtbookToDaisy3.Parameters["include-tts-log"].ParameterValue == true)
            {
                dtbookToDaisy3.Parameters["tts-log"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "tts-log")).FullName;
            }

            dtbookToDaisy3.Parameters["output"].ParameterValue = daisy3TempDir.FullName;
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting " + dtbookToDaisy3.Parameters["input"].ParameterValue + " dtbook XML to DAISY3 in " + dtbookToDaisy3.Parameters["output"].ParameterValue));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting dtbook XML to DAISY3... "));
#endif
            
            dtbookToDaisy3.ExecuteScript(dtbookToDaisy3.Parameters["input"].ParameterValue.ToString());
            foreach (var kv in this._parameters) {
                if (daisy3ToMp3.Parameters.ContainsKey(kv.Key) && kv.Key != "output") // avoid copy daisy3 to mp3 output as it is not the same named option
                {
                    daisy3ToMp3.Parameters[kv.Key] = kv.Value;
                }
            }
            try
            {
                daisy3ToMp3.Parameters["input"].ParameterValue = Directory.GetFiles(daisy3TempDir.FullName, "*.opf", SearchOption.AllDirectories)[0];
            }
            catch
            {
                throw new FileNotFoundException("Could not find result of cleaning process in result folder", tempDir.FullName);
            }
            daisy3ToMp3.Parameters["output"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "MP3 files")).FullName;
            daisy3ToMp3.ExecuteScript((string)daisy3ToMp3.Parameters["input"].ParameterValue, true);
        }
    }
}
