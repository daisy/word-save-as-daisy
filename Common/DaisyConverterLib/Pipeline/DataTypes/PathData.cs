using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class PathData : ParameterDataBase
    {
        private string path;
        public enum InputOrOutput { input, output };
        private InputOrOutput inputOrInput;
        public bool IsInput { get { return inputOrInput == InputOrOutput.input; } }
        public bool IsOutput { get { return inputOrInput == InputOrOutput.output; } }

        public enum FileOrDirectory { File, Directory }
        private FileOrDirectory fileOrDirectory;
        public bool IsFile { get { return fileOrDirectory == FileOrDirectory.File; } }
        public bool IsDirectory { get { return fileOrDirectory == FileOrDirectory.Directory; } }


        public string FileExtension { get; private set; }


        public PathData(
            InputOrOutput portType = InputOrOutput.output,
            FileOrDirectory pathType = FileOrDirectory.Directory,
            string mimeType = "",
            string initialPath = ""
         ) : base()
        {
            path = initialPath;
            fileOrDirectory = pathType;
            inputOrInput = portType;
            FileExtension = GetFileExtensionFromMimeType(mimeType);
        }


        public override object Value
        {
            get {
                return path;
                //if (ShouldExists(path)) {
                //    return path;
                //} else throw new System.Exception(path + " does not exist");
            }
            set
            {
                path = (string)value;
                //if (ShouldExists((string)value))
                //{
                //    path = (string)value;
                //}
                //else throw new System.Exception("No_Path");
            }
        }
        /// <summary>
        ///  Does the path exists
        /// </summary>
        /// <param name="PathValue"></param>
        /// <returns></returns>
        public bool ShouldExists(string PathValue)
        {
            if (IsInput) {
                if (IsFile && File.Exists(PathValue)) {
                    return true;
                } else if (Directory.Exists(PathValue)) {
                    return true;
                }
            } else if (IsOutput) {
                return true;
            }
            return false;
        }


        public bool Create(string pathValue)
        {
            if (inputOrInput == InputOrOutput.output)
            {
                if (fileOrDirectory == FileOrDirectory.File)
                {
                    File.CreateText(pathValue);
                    return true;
                }
                else if (fileOrDirectory == FileOrDirectory.Directory)
                {
                    Directory.CreateDirectory(pathValue);
                    return true;
                }
                return false;
            }
            else
            {
                throw new System.Exception("CannotCreate_InputPath");
            }
        }

        protected string GetFileExtensionFromMimeType(string mimetype)
        {
            switch (mimetype)
            {
                case "application/epub+zip":
                    return ".epub";
                case "application/oebps-package+xml":
                    return ".opf";
                case "application/xml":
                    return ".xml";
                default:
                    return string.Empty;
            }
        }

    }


}
