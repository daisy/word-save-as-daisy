using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        public static string DTBookToODTTemplate
        {
            get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\dtbook-to-odt.ott"; }
        }

        public static string EmbeddedEngineProperties
        {
            get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\daisy-pipeline\settable-properties.xml"; }
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




        private static readonly StringValidator _dTBookXMLFileNameFormat = new StringValidator {
            AuthorisationPattern = new Regex(@"^[^,]+$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Any commas (,) present in the file name should be removed, or they will be replaced by underscores automatically."
        };
		public static StringValidator DTBookXMLFileNameFormat { get => _dTBookXMLFileNameFormat; }

        
        private static readonly StringValidator _DAISYFileNameFormat = new StringValidator {
            AuthorisationPattern = new Regex(@"^[a-zA-Z0-9_\-\.]+\.docx$"),
            UnauthorizedValueMessage = "Your document file name contains unauthorized characters, that may be automatically replaced by underscores.\r\n" +
                        "Only Alphanumerical letters (a-z, A-Z, 0-9), hyphens (-), dots (.) " +
                            "and underscores (_) are allowed in DAISY file names." +
                        "\r\nAny other characters (including spaces) will be replaced automaticaly by underscores."
        };
		public static StringValidator DAISYFileNameFormat { get => _DAISYFileNameFormat; }

        /// <summary>
        /// Looks for the Ace by DAISY executable in its usual installation locations
        /// and in the PATH environment variable.
        /// </summary>
        /// <returns>The full path of the Ace by DAISY executable, or null if it is not installed</returns>
        public static string FindAceByDaisyExecutable()
        {
            const string aceExecutableName = "Ace by DAISY.exe";
            List<string> candidates = new List<string>();
            foreach (string envVar in new[] { "LocalAppData", "ProgramFiles", "ProgramFiles(x86)", "ProgramW6432" })
            {
                string root = Environment.GetEnvironmentVariable(envVar);
                if (string.IsNullOrEmpty(root))
                {
                    continue;
                }
                // Per-user installation (default install location of the Ace by DAISY setup)
                candidates.Add(Path.Combine(root, "Programs", "Ace by DAISY", aceExecutableName));
                // Per-machine installation
                candidates.Add(Path.Combine(root, "Ace by DAISY", aceExecutableName));
            }
            string pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                foreach (string dir in pathEnv.Split(';'))
                {
                    if (!string.IsNullOrWhiteSpace(dir))
                    {
                        candidates.Add(Path.Combine(dir.Trim(), aceExecutableName));
                    }
                }
            }
            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            return null;
        }


        /// <summary>
        /// Launches Ace by DAISY with the epub file as argument.
        /// As Ace by DAISY may not run correctly on first launch, it is first started without
        /// any argument, then relaunched with the epub file after a 2 seconds delay.
        /// </summary>
        /// <param name="acePath">Path of the Ace by DAISY executable</param>
        /// <param name="epubFile">Path of the epub file</param>
        public static void LaunchAceByDaisyOnEpub(string acePath, string epubFile)
        {
            try
            {
                AddinLogger.Info($"Launching Ace by DAISY on \"{epubFile}\"");
                // Launch in a background task to avoid freezing the Word UI during the 2 seconds delay
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        // First launch without argument: Ace by DAISY may not run correctly on first launch
                        Process.Start(new ProcessStartInfo(acePath) { UseShellExecute = true });
                        System.Threading.Thread.Sleep(2000);
                        // Relaunch with the epub file as argument
                        Process.Start(new ProcessStartInfo(acePath, $"\"{epubFile}\"") { UseShellExecute = true });
                    }
                    catch (Exception e)
                    {
                        AddinLogger.Error("An error occured while launching Ace by DAISY", e);
                    }
                });
            }
            catch (Exception e)
            {
                AddinLogger.Error("An error occured while looking for an epub file to open in Ace by DAISY", e);
            }
        }

    }
}
