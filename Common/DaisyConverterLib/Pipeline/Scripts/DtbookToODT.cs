using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class DtbookToODT : Script
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;
        public DtbookToODT(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "dtbook-to-odt";
            this.niceName = "DTBook XML to ODT";
            _parameters.Add("input", new ScriptParameter(
                "source",
                "DTBook XML file",
                new PathData(
                    PathData.InputOrOutput.input,
                    PathData.FileOrDirectory.File,
                    "application/xml"
                ),
                true,
                "The DTBook file to be converted"
            ));
            _parameters.Add("output", new ScriptParameter(
                "result",
                "Export directory",
                new PathData(
                    PathData.InputOrOutput.output,
                    PathData.FileOrDirectory.Directory
                ),
                true,
                "The directory where the resulting ODT file will be stored."
            ));
            _parameters.Add("template", new ScriptParameter(
                "template",
                "Template OTT",
                new PathData(
                    PathData.InputOrOutput.input,
                    PathData.FileOrDirectory.File,
                    "application/vnd.oasis.opendocument.text-template",
                    ConverterHelper.DTBookToODTTemplate
                ),
                false,
                "OpenOffice template file (.ott) that contains the style definitions."
            ));
            _parameters.Add("asciimath", new ScriptParameter(
                "asciimath",
                "ASCIIMath handling",
                new EnumData(
                    new Dictionary<string, object>(){
                        {"Keep (as plain text)", "ASCIIMATH"},
                        {"Convert to MathML", "MATHML"},
                    }, "ASCIIMATH"),
                true,
                "How to render ASCIIMath-encoded formulas.\r\n\r\n" +
                "Mathematical content in the DTBook may be encoded as [MathML](https://en.wikipedia.org/wiki/MathML)\r\n" +
                "elements, or as plain [ASCIIMath](https://asciimath.org/) text enclosed in `span` elements with a\r\n" +
                "\"asciimath\" class. MathML elements are converted to formula objects. ASCIIMath can either be treated\r\n" +
                "the same as MathML, or included as-is."
            ));

            _parameters.Add("images", new ScriptParameter(
                "images",
                "Images handling",
                new EnumData(
                    new Dictionary<string, object>(){
                        {"Embed images in the OpenDocument package", "EMBED"},
                        {"Link to external image files", "LINK"},
                    }, "EMBED"),
                false,
                "How to render images."
            ));

            _parameters.Add("page-number", new ScriptParameter(
                "page-numbers",
                "Page numbers",
                new BoolData(true),
                false,
                "Whether to show page numbers or not."
            ));

            _parameters.Add("page-numbers-float", new ScriptParameter(
                "page-numbers-float",
                "Float page numbers",
                new BoolData(true),
                false,
                "Try to float page numbers to an appropriate place as opposed to exactly following print."
            ));
            _parameters.Add("image-dpi", new ScriptParameter(
                "image-dpi",
                "Image resolution.",
                ConverterSettings.ImageResamplingChoice.DataType(),
                false,
                "Resolution of images in DPI.",
                false
            ));

        }
        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {
            throw new NotImplementedException();
        }
    }
}
