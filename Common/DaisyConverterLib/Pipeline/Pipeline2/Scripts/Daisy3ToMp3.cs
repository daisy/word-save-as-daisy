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
