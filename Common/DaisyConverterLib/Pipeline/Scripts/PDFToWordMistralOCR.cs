using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class PDFToWordMistralOCR : Script
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public PDFToWordMistralOCR(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "pdf-to-word-mistral-ocr";
            this.niceName = "PDF to Word using Mistral OCR (experimental)";

            _parameters.Add("input", new ScriptParameter(
                "source",
                "PDF file",
                new PathData(
                    PathData.InputOrOutput.input,
                    PathData.FileOrDirectory.File,
                    "application/pdf"
                ),
                true,
                "The PDF file to be converted"
            ));
            _parameters.Add("output", new ScriptParameter(
                "result",
                "Export directory",
                new PathData(
                    PathData.InputOrOutput.output,
                    PathData.FileOrDirectory.Directory
                ),
                true,
                "The directory where the resulting Word file will be stored."
            ));
            _parameters.Add("model", new ScriptParameter(
                "model",
                "Model version",
                new EnumData(new Dictionary<string, object>()
                {
                    { "mistral-ocr-2512",   "mistral-ocr-2512"},
                    { "mistral-ocr-latest", "mistral-ocr-latest"},
                    { "mistral-ocr-2505",   "mistral-ocr-2505"}
                }, "mistral-ocr-2512"),
                false,
                "The Mistral OCR model to be used."
            ));
            _parameters.Add("include-html", new ScriptParameter(
                "include-html",
                "Include HTML",
                new BoolData(false),
                false,
                "Whether or not to keep the intermediary HTML file set (for debugging)."
            ));

        }
        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}

