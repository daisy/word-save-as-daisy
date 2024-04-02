using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class CleanedDtbookToDaisy3 : Script {

        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        Script dtbookCleaner;
        Script dtbookToDaisy3;

        public CleanedDtbookToDaisy3(IConversionEventsHandler e): base(e) {
            this.niceName = "Export to DAISY3";
            dtbookCleaner = new DtbookCleaner(e);
            dtbookToDaisy3 = new DtbookToDaisy3(e);
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
                        "output-dir",
                        "Daisy3 output directory",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The resulting DAISY 3 publication."
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
                        true,
                        false,
                        "Whether to use a speech synthesizer to produce audio files."
                    )
                },
                {"with-text", new ScriptParameter(
                        "with-text",
                        "With text",
                        new BoolDataType(),
                        true,
                        false,
                        "Includes DTBook in output, as opposed to audio only."
                    )
                },
                {"word-detection", new ScriptParameter(
                        "word-detection",
                        "Apply word detection",
                        new BoolDataType(),
                        false,
                        false,
                        "Whether to detect and mark up words with <w> tags.\r\n" +
                        "\r\n" +
                        "By default word detection is performed but an option is provided to disable it because some DAISY 3\r\n" +
                        "reading systems can't handle the word tags."
                    )
                },
            };
        }

        public override void ExecuteScript(string inputPath, bool isQuite) {

            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["output"].ParameterValue.ToString(),
                string.Format(
                    "{0}_DAISY3_{1}",
                    Path.GetFileNameWithoutExtension(inputPath),
                    DateTime.Now.ToString("yyyyMMddHHmmssffff")
                )
            ) ;
            // Remove and recreate result folder
            // Since the DaisyToEpub3 requires output folder to be empty
            if (Directory.Exists(finalOutput)) {
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
            dtbookCleaner.ExecuteScript(inputPath,true);

            // transfer parameters value
            foreach(var kv in this._parameters)
            {
                if (dtbookToDaisy3.Parameters.ContainsKey(kv.Key))
                {
                    dtbookToDaisy3.Parameters[kv.Key] = kv.Value;
                }
            }
            // rebind input and output
            try
            {
                dtbookToDaisy3.Parameters["input"].ParameterValue = Directory.GetFiles(tempDir.FullName, "*.xml", SearchOption.AllDirectories)[0];
            } catch
            {
                throw new FileNotFoundException("Could not find result of cleaning process in result folder", tempDir.FullName);
            }
            
            dtbookToDaisy3.Parameters["output"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "DAISY3")).FullName;
            if ((bool)dtbookToDaisy3.Parameters["include-tts-log"].ParameterValue == true)
            {
                dtbookToDaisy3.Parameters["tts-log"].ParameterValue = Directory.CreateDirectory(Path.Combine(finalOutput, "tts-log")).FullName;
            }
#if DEBUG
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting " + dtbookToDaisy3.Parameters["input"].ParameterValue + " dtbook XML to DAISY3 in " + dtbookToDaisy3.Parameters["output"].ParameterValue));
#else
            this.EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Converting dtbook XML to DAISY3... "));
#endif

            dtbookToDaisy3.ExecuteScript(dtbookToDaisy3.Parameters["input"].ParameterValue.ToString());
        }
    }
}
