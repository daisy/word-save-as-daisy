using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {
    public class ConverterHelper {

		/// <summary>
		/// Gets path to pipeline root directory.
		/// </summary>
		public static string PipelinePath {
			get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms"; }
		}

		/// <summary>
		/// Indicates if pipeline exists.
		/// </summary>
		/// <returns></returns>
		public static bool PipelineIsInstalled() {
			return Directory.Exists(PipelinePath);
		}

		/// <summary>
		/// Gets path to the addin directory in AppData.
		/// </summary>
		public static string AppDataSaveAsDAISYDirectory {
			get { return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY"; }
		}

		public static bool documentIsOpen(string documentPath) {
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




		public static FilenameValidator DTBookXMLFileNameFormat = new FilenameValidator {
			AuthorisationPattern = new Regex(@"^[^,]+$"),
			UnauthorizedNameMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
						"Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically."
		};

		public static FilenameValidator DAISYFileNameFormat = new FilenameValidator {
			AuthorisationPattern = new Regex(@"^[a-zA-Z0-9_\-\.]+\.docx$"),
			UnauthorizedNameMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
						"Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
							"and underscores (_) are allowed in DAISY file names." +
						"\r\nAny other characters (including spaces) will be replaced automaticaly by underscores."
		};

	}
}
