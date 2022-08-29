using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts {
    public class DtbookToEpub3 : Script {

        Script dtbookToDaisy;
        Script DaisyToEpub3;

        public DtbookToEpub3() {
            dtbookToDaisy = Pipeline1.Instance.getScript("default");
            DaisyToEpub3 = new Daisy3ToEPUB3();
            // TODO : for now we consider the 3 global steps of the progression but some granularity within
            // scripts could be taked in account
            StepsCount = 4;
            // use default script parameters to not change the form for now
            this._parameters = dtbookToDaisy.Parameters;
        }

        public override void ExecuteScript(string inputPath, bool isQuite) {

            // Create a directory using the document name
            string finalOutput = Path.Combine(
                Parameters["outputPath"].ParameterValue.ToString(),
                Path.GetFileNameWithoutExtension(inputPath)
                ) ;
            // Remove and recreate result folder
            // Since the DaisyToEpub3 requires output folder to be empty
            if (Directory.Exists(finalOutput)) {
                Directory.Delete(finalOutput, true);
            }
            Directory.CreateDirectory(finalOutput);

            DirectoryInfo tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()));

            if (this.OnPipelineProgress != null) {
                OnPipelineProgress("Converting dtbook to DAISY3");
                dtbookToDaisy.setPipelineProgressListener(OnPipelineProgress);
            }
            if (this.OnPipelineOutput != null) {
                dtbookToDaisy.setPipelineOutputListener(OnPipelineOutput);
            }
            if (OnPipelineError != null) {
                dtbookToDaisy.setPipelineErrorListener(OnPipelineError);
            }
            dtbookToDaisy.Parameters["outputPath"].ParameterValue = tempDir.FullName;
            dtbookToDaisy.ExecuteScript(inputPath,true);

            
            DaisyToEpub3.Parameters["input"].ParameterValue = Directory.GetFiles(tempDir.FullName, "*.opf")[0];
            
            DaisyToEpub3.Parameters["output"].ParameterValue = finalOutput;
            if (this.OnPipelineProgress != null) {
                OnPipelineProgress("Converting DAISY3 to EPUB3");
                DaisyToEpub3.setPipelineProgressListener(OnPipelineProgress);
            }
            if (this.OnPipelineOutput != null) {
                DaisyToEpub3.setPipelineOutputListener(OnPipelineOutput);
            }
            if(OnPipelineError != null) {
                DaisyToEpub3.setPipelineErrorListener(OnPipelineError);
            }
            DaisyToEpub3.ExecuteScript(DaisyToEpub3.Parameters["input"].ParameterValue.ToString());

            if (File.Exists(Path.Combine(finalOutput, "result.epub"))) {
                if (this.OnPipelineProgress != null) {
                    OnPipelineProgress("Moving result.epub to " + Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub"));
                    //DaisyToEpub3.setPipelineProgressListener(OnPipelineProgress);
                }
                File.Move(
                    Path.Combine(finalOutput, "result.epub"),
                    Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub")
                );
                if (this.OnPipelineProgress != null) {
                    OnPipelineProgress("Successfully converted the document to " + Path.Combine(finalOutput, Path.GetFileNameWithoutExtension(inputPath) + ".epub"));
                    //DaisyToEpub3.setPipelineProgressListener(OnPipelineProgress);
                }
            }
        }
    }
}
