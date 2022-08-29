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
        private string m_Path;

        public enum InputOrOutput { input, output } ;
        private InputOrOutput m_InputOrOutput;

        public enum FileOrDirectory { File, Directory }
        private FileOrDirectory m_FileOrDirectory;

		public string FileExtension { get; private set; }

        public PathDataType(ScriptParameter p, XmlNode DataTypeNode) : base(p)
        {
            m_Path = p.ParameterValue.ToString();
            XmlNode ChildNode = DataTypeNode.FirstChild;

        	if (ChildNode.Name == "file")
        	{
        		m_FileOrDirectory = FileOrDirectory.File;
                XmlAttribute fileTypeAttribute = ChildNode.Attributes["mime"];
                string attributeValue = fileTypeAttribute == null ? string.Empty : fileTypeAttribute.Value;
                FileExtension = GetFileExtensionFromMimeType(attributeValue);
            }
        	else
        	{
        		m_FileOrDirectory = FileOrDirectory.Directory;
        	}

        	m_InputOrOutput = ChildNode.Attributes.GetNamedItem("type").Value == "input" ? InputOrOutput.input : InputOrOutput.output;
        }

        public PathDataType(
            InputOrOutput portType = InputOrOutput.output,
            FileOrDirectory pathType = FileOrDirectory.Directory,
            string mimeType = "",
            string initialPath = ""
         ) : base() {
            m_Path = initialPath;
            m_FileOrDirectory = pathType;
            m_InputOrOutput = portType;
            FileExtension = GetFileExtensionFromMimeType(mimeType);
        }

        /// <summary>
        ///  The path is of a file or a directory
        /// </summary>
        public FileOrDirectory IsFileOrDirectory { get { return m_FileOrDirectory; } }

        /// <summary>
        /// Is it a path for input file/directory or for output file/directory
        /// </summary>
        public InputOrOutput isInputOrOutput { get { return m_InputOrOutput; } }

        public override object Value
        {
            get { return m_Path; }
            set
            {
                if (Exists((string)value))
                {
                    m_Path = (string)value;
                    UpdateScript(m_Path);
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
            if (m_FileOrDirectory == FileOrDirectory.File)
            {
                if (File.Exists(PathValue))
                    return true;
            }
            else if (m_FileOrDirectory == FileOrDirectory.Directory)
            {
                if (Directory.Exists(PathValue))
                    return true;
            }
            return false;
        }

        private bool UpdateScript(string Val)
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
            if (m_InputOrOutput == InputOrOutput.output)
            {
                if (m_FileOrDirectory == FileOrDirectory.File)
                {
                    File.CreateText(pathValue);
                    return true;
                }
                else if (m_FileOrDirectory == FileOrDirectory.Directory)
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
