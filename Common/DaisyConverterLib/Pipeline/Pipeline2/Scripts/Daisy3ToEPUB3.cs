﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts {
    public class Daisy3ToEPUB3 : Pipeline2Script {

        public Daisy3ToEPUB3() {
            this.nameOrURI = "daisy3-to-epub3";
            this.niceName = "DAISY3 to EPUB3";
            _parameters = new Dictionary<string, ScriptParameter>();
            _parameters.Add(
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
                    "The package file of the input DTB."
                )
            );
            _parameters.Add(
                "output",
                new ScriptParameter(
                    "output-dir",
                    "EPUB output directory",
                    new PathDataType(
                        PathDataType.InputOrOutput.output,
                        PathDataType.FileOrDirectory.Directory
                    ),
                    "",
                    true,
                    "The produced EPUB."
                )
            );
            // I'm not sure but the temp-dir option is reported unkown
            // when trying to launch the script, while i found it in the xpl script
            //_parameters.Add(
            //    "temp",
            //    new ScriptParameter(
            //        "temp-dir",
            //        "temp directory",
            //        new PathDataType(
            //            PathDataType.InputOrOutput.output,
            //            PathDataType.FileOrDirectory.Directory
            //        ),
            //        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(),Path.GetRandomFileName())).FullName,
            //        true,
            //        "Temp directory to use"
            //    )
            //);
            _parameters.Add(
                "mediaoverlays",
                new ScriptParameter(
                    "mediaoverlays",
                    "Include media overlays",
                    new BoolDataType(true),
                    "true",
                    false,
                    "Whether or not to include media overlays and associated audio files (true or false)"
                )
            );
            _parameters.Add(
                "assert-valid",
                new ScriptParameter(
                    "assert-valid",
                    "Assert validity",
                    new BoolDataType(false),
                    "false",
                    false
                )
            );
        }
    }
}
