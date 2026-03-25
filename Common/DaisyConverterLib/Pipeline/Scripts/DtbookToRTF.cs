using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Scripts
{
    public class DtbookToRTF : Script
    {
        public DtbookToRTF(IConversionEventsHandler e)
            : base(e)
        {
            this.name = "dtbook-to-rtf";
            this.niceName = "DTBook XML to RTF";
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
                "The directory where the resulting RTF file will be stored."
            ));
            _parameters.Add("include-table-of-content", new ScriptParameter(
                "include-table-of-content",
                "Include table of content",
                new BoolData(false),
                false,
                "A boolean indicating if a TOC should be generated."
            ));

            _parameters.Add("include-page-number", new ScriptParameter(
                "include-page-number",
                "Include page number",
                new BoolData(false),
                false,
                "A boolean indicating if a page numbers should be inserted."
            ));

        }
        public override string searchInputFromDirectory(DirectoryInfo inputDirectory)
        {

            throw new NotImplementedException();
        }
    }
}
