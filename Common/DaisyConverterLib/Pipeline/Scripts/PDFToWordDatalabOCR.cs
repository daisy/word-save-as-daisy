using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;


namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class PDFToWordDatalabOCR : Script
    {
        //private static readonly ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public PDFToWordDatalabOCR(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "pdf-to-word-datalab";
            this.niceName = "PDF to Word using Datalab (experimental)";
            this.description = "Transforms a PDF into a Microsoft Office Word (.docx) document, powered by Datalab.\r\nThis script is an early development phase and has not been tested with a wide range of input documents yet.";
            _parameters.Add("input", new ScriptParameter(
                "source",
                "PDF file",
                new PathData(
                    PathData.InputOrOutput.input,
                    PathData.FileOrDirectory.File,
                    "application/pdf"
                ),
                true,
                "The PDF you want to convert."
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
                    { "fast",   "fast"},
                    { "balanced", "balanced"},
                    { "accurate",   "accurate"}
                }, "balanced"),
                false,
                "The Datalab model to be used."
            ));
            _parameters.Add("include-html", new ScriptParameter(
                "include-html",
                "Include HTML",
                new BoolData(false),
                false,
                "Whether or not to keep the intermediary HTML file set (for debugging)."
            ));

        }
        public override string SearchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }

        // Force using the embedded engine for PDF to word conversion
        protected override Runner GetRunner()
        {
            return EmbeddedRunner.GetInstance(EventsHandler);
        }
    }
}

