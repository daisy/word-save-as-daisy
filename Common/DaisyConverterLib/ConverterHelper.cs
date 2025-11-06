using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {
    public static class ConverterHelper {

		/// <summary>
		/// Get the pipeline 2 root directory
		/// </summary>
		public static string EmbeddedEnginePath {
			get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\daisy-pipeline"; }
		}

		/// <summary>
		/// Indicates if pipeline exists.
		/// </summary>
		/// <returns></returns>
		public static bool PipelineIsInstalled() {
			return Directory.Exists(EmbeddedEnginePath);
		}
        public static string EmbeddedEngineLauncherPath {
            get { 
				return Path.Combine(
					Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
					"daisy-pipeline", "bin", "pipeline2.bat"
				);
			}
        }
        /// <summary>
        /// Gets path to the addin directory in AppData.
        /// </summary>
        public static string AppDataSaveAsDAISYDirectory {
			get { return Directory.CreateDirectory(
				Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
					"SaveAsDAISY"
				)).FullName; }
		}

		public static readonly string PipelineAppReleaseURL = "https://github.com/daisy/pipeline-ui/releases/latest";

        public static string PipelineAppPath {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs", "pipeline-ui", "DAISY Pipeline.exe"
                );
            }
        }

        public static string PipelineAppDataPath {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "pipeline-ui"
                );
            }
        }
        public static string PipelineAppTTSConfigPath{
            get
            {
                return Path.Combine(
                    PipelineAppDataPath, "ttsConfig.xml"
                );
            }
        }

        public static bool PipelineAppIsInstalled()
        {
            return File.Exists(PipelineAppPath);
        }



        public static bool DocumentIsOpen(string documentPath) {
			try {
				Package pack;
				pack = Package.Open(documentPath, FileMode.Open, FileAccess.ReadWrite);
				pack.Close();
			} catch {
				return true;
			}
			return false;
		}

		public static string GetTempPath(string fileName, string targetExtension) {
			string folderName;
			string path;
			do {
				folderName = Path.GetRandomFileName();
				path = Path.Combine(Path.GetTempPath(), folderName);
			}
			while (Directory.Exists(path));

			Directory.CreateDirectory(path);
			return Path.Combine(path, Path.GetFileNameWithoutExtension(fileName) + targetExtension);
		}

		/// <summary>
		/// Function to generate Random ID
		/// </summary>
		/// <returns></returns>
		static public long GenerateId() {
			byte[] buffer = Guid.NewGuid().ToByteArray();
			return BitConverter.ToInt64(buffer, 0);
		}




        private static StringValidator _dTBookXMLFileNameFormat = new StringValidator {
            AuthorisationPattern = new Regex(@"^[^,]+$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically."
        };
		public static StringValidator DTBookXMLFileNameFormat { get => _dTBookXMLFileNameFormat; }

        
        private static StringValidator _DAISYFileNameFormat = new StringValidator {
            AuthorisationPattern = new Regex(@"^[a-zA-Z0-9_\-\.]+\.docx$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
                            "and underscores (_) are allowed in DAISY file names." +
                        "\r\nAny other characters (including spaces) will be replaced automaticaly by underscores."
        };
		public static StringValidator DAISYFileNameFormat { get => _DAISYFileNameFormat; }



	}
}
