using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class PathDataType : ParameterDataType
    {
        private string path;

        //TODO : This can be replaced by a simple boolean IsOutput or IsInput

        public enum InputOrOutput { input, output };
        private InputOrOutput inputOrInput;
        public bool IsInput { get { return inputOrInput == InputOrOutput.input; } }
        public bool IsOutput { get { return inputOrInput == InputOrOutput.output; } }

        // TODO : this can be replace by a simple IsFile or IsDirectory boolean

        public enum FileOrDirectory { File, Directory }
        private FileOrDirectory fileOrDirectory;
        public bool IsFile { get { return fileOrDirectory == FileOrDirectory.File; } }
        public bool IsDirectory { get { return fileOrDirectory == FileOrDirectory.Directory; } }



        public string FileExtension { get; private set; }

        public PathDataType(ScriptParameter p, XmlNode DataTypeNode) : base(p)
        {
            path = p.ParameterValue.ToString();
            XmlNode ChildNode = DataTypeNode.FirstChild;

            if (ChildNode.Name == "file")
            {
                fileOrDirectory = FileOrDirectory.File;
                XmlAttribute fileTypeAttribute = ChildNode.Attributes["mime"];
                string attributeValue = fileTypeAttribute == null ? string.Empty : fileTypeAttribute.Value;
                FileExtension = GetFileExtensionFromMimeType(attributeValue);
            }
            else
            {
                fileOrDirectory = FileOrDirectory.Directory;
            }

            inputOrInput = ChildNode.Attributes.GetNamedItem("type").Value == "input" ? InputOrOutput.input : InputOrOutput.output;
        }

        public PathDataType(
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
            get { return path; }
            set
            {
                if (Exists((string)value))
                {
                    path = (string)value;
                    UpdateScript(path);
                }
                else throw new System.Exception("No_Path");
            }
        }
        /// <summary>
        ///  Does the path exists
        /// </summary>
        /// <param name="PathValue"></param>
        /// <returns></returns>
        public bool Exists(string PathValue)
        {
            if (IsFile)
            {
                if (File.Exists(PathValue))
                    return true;
            }
            else
            {
                if (Directory.Exists(PathValue))
                    return true;
            }
            return false;
        }

        private new bool UpdateScript(string Val)
        {
            if (Val != null)
            {
                m_Parameter.ParameterValue = Val;
                return true;
            }
            else
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
                default:
                    return string.Empty;
            }
        }

    }


}
