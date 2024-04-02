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
                  "player-type",
                  new ScriptParameter(
                    "player-type",
                    "The type of MegaVoice Envoy device.",
                    new EnumDataType(
                        new Dictionary<string,object>(){
                            {"Type 1", "1"},
                            {"Type 2", "2"},
                        }, "Type 1"),
                    "1",
                    false,
                    "Possible values: \r\n" +
                    " - Type 1: This produces a folder structure that is four levels deep." +
                    " At the top level there can be up to 8 folders." +
                    " Each of these folders can have up to 20 sub-folders." +
                    " The sub-folders can have up to 999 sub-sub-folders, each of which can contain up to 999 MP3 files.\r\n" +
                    " - Type 2: This produces a folder structure that is two levels deep. On the top level there can be up to 999 folders, and each of these folders can have up to 999 MP3 files."
                  )
                },
                {
                  "book-folder-level",
                  new ScriptParameter(
                    "book-folder-level",
                    "Book folder level",
                    new IntegerDataType(0,3,1),
                    1,
                    false,
                    "The folder level corresponding with the book, expressed as a 0-based number.\r\n" +
                    "Value `0` means that top-level folders correspond with top-level sections of the book.\r\n" +
                    "Value `1` means that the book is contained in a single top-level folder (if it is not too big) with" +
                    " sub - folders that correspond with top-level sections of the book." +
                    " Must be a non-negative integer value, less than or equal to 3 for type 1 players, and less than or" +
                    " equal to 1 for type 2 players."
                  )
                }
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
