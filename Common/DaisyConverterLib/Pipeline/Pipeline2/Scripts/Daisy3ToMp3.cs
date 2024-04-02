using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts {
    public class Daisy3ToMp3 : Pipeline2Script {

        public Daisy3ToMp3(IConversionEventsHandler eventsHandler) : base(eventsHandler) {
            this.name = "daisy3-to-mp3";
            _parameters = new Dictionary<string, ScriptParameter>
            {
                {
                    "input",
                    new ScriptParameter(
                        "source",
                        "OPF",
                        new PathDataType(
                            PathDataType.InputOrOutput.input,
                            PathDataType.FileOrDirectory.File,
                            "application/oebps-package+xml"
                        ),
                        "",
                        true,
                        "The package file of the input DAISY3."
                    )
                },
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
                    " - Type 1: This produces a folder structure that is four levels deep.\r\n" +
                    " At the top level there can be up to 8 folders." +
                    " Each of these folders can have up to 20 sub-folders." +
                    " The sub-folders can have up to 999 sub-sub-folders, each of which can contain up to 999 MP3 files.\r\n" +
                    " - Type 2: This produces a folder structure that is two levels deep.\r\n" +
                    " On the top level there can be up to 999 folders, and each of these folders can have up to 999 MP3 files."
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
                    "The folder level corresponding with the book, expressed as a 0-based number.\r\n\r\n" +
                    "Value `0` means that top-level folders correspond with top-level sections of the book.\r\n" +
                    "Value `1` means that the book is contained in a single top-level folder (if it is not too big) with sub" +
                    "- folders that correspond with top-level sections of the book.\r\n" +
                    "Must be a non-negative integer value, less than or equal to 3 for type 1 players, and less than or equal to 1 for type 2 players."
                  )
                },
                {
                    "output",
                    new ScriptParameter(
                        "result",
                        "MP3 Files",
                        new PathDataType(
                            PathDataType.InputOrOutput.output,
                            PathDataType.FileOrDirectory.Directory
                        ),
                        "",
                        true,
                        "The produced folder structure with MP3 files."
                    )
                },
            };
        }
    }
}
