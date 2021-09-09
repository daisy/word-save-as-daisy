using System;
using System.Diagnostics;
using System.IO;
using NUnit.Framework;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Collections.Generic;
using Daisy.SaveAsDAISY.Addins.Word2007;
using Extensibility;
using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;

namespace RegressionTests {

	public class Tests {
		protected Application _word;
		protected Document _document;
		public const string OfficeVersion = "14";

		protected object falseValue = false;
		protected object trueValue = true;
		protected object missing = Type.Missing;

		// Default folders to be used for tests

		public DirectoryInfo testfilesDirectory;
		public DirectoryInfo inputDirectory;
		public DirectoryInfo expectedDirectory;
		public DirectoryInfo outputDirectory;
		public DirectoryInfo reportsDirectory;


		[OneTimeSetUp]
		public void FixtureSetUp() {
			// VS 2019 - Change working directory as the default one is now in a temp directory of APPDATA
			// Relocating the working directory to the project directory from the assembly location 
			// (location is alledgedly <project_folder>/bin/<configuration>/file.dll 
			// see the output path within the build tab of the project properties)
			string location = Assembly.GetExecutingAssembly().Location;
			Directory.SetCurrentDirectory(
				Directory.GetParent(
					Directory.GetParent(
						Directory.GetParent(
							Assembly.GetExecutingAssembly().Location
						).FullName
					).FullName
				).FullName
			);

			testfilesDirectory = new DirectoryInfo(@"testfiles");
			inputDirectory = new DirectoryInfo(@"testfiles/input");
			expectedDirectory = new DirectoryInfo(@"testfiles/expected");
			outputDirectory = new DirectoryInfo(@"testfiles/output");
			reportsDirectory = new DirectoryInfo(@"testfiles/reports");


			try {
				if (Directory.Exists(outputDirectory.FullName))
					Directory.Delete(outputDirectory.FullName, true);

				Directory.CreateDirectory(outputDirectory.FullName);
				
			} catch (Exception ex) {
				Console.WriteLine("Some error during output folders preparation : " + ex.Message);
			}
		}

		[OneTimeTearDown]
		public void FixtureTearDown() {
			if (Directory.Exists(outputDirectory.FullName)) {
				try {
					Directory.Delete(outputDirectory.FullName, true);
				} catch (Exception ex) {
					Console.WriteLine("Can'not delete output folder : " + ex.Message);
				}

			}
			CloseWord();
		}

		[TearDown]
		public void TearDown() {
			CloseWord();
		}

		[Test]
		public void batchRegressionTests() {
			List<string> documents = getDocumentList(inputDirectory);
			

			_word = new Microsoft.Office.Interop.Word.Application();
			_word.Visible = false;
			//_word.Activate();

			//Connect connect = new Connect();

			//Array array = new object[0];
			//connect.OnConnection(_word, ext_ConnectMode.ext_cm_External, null, ref array);

			string testReportPath = Path.Combine(
				reportsDirectory.FullName,
				"regressionReport-" + DateTime.Now.ToBinary().ToString() + ".log"
			);
			StreamWriter testReport = new StreamWriter(File.OpenWrite(testReportPath));
			

			foreach (string doc in documents) {
				Console.WriteLine("Converting " + doc);
				object fileName = doc;
				//_document = _word.Documents.Open(ref fileName, ref missing, ref falseValue, ref trueValue, missing, missing, missing, missing, missing, missing, missing, falseValue); ;
				//_document.Activate();
				try {
					// Do a conversion and put it in a subfolder in output
					string conversionOutput = Path.Combine(
						outputDirectory.FullName,
						Path.GetFileNameWithoutExtension(doc)
					);
					DirectoryInfo expectedOutput = new DirectoryInfo(
						Path.Combine(
							expectedDirectory.FullName,
							Path.GetFileNameWithoutExtension(doc)
						)
					);

					// Note : for now i reuse the operation used in Converter.convert method
					string outputName = Path.GetFileNameWithoutExtension(doc).Replace(",", "_") + ".xml";

					if (Directory.Exists(conversionOutput)) {
						Directory.Delete(conversionOutput, true);
                    }
					Directory.CreateDirectory(conversionOutput);

					WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
					ConversionParameters parameters = new ConversionParameters()
						.withParameter("Visible", false)
						.withParameter("OutputFile", Path.Combine(conversionOutput, outputName))
						.withParameter("TrackChangesFlag", "NoTrack")
						.withParameter("Version", OfficeVersion)
						.withParameter("MasterSubFlag", "No");

					IDocumentPreprocessor preprocess = new DocumentPreprocessor(_word);
					IConversionEventsHandler eventsHandler = null;
					eventsHandler = new SilentEventsHandler();
					Converter converter = new Converter(preprocess, documentConverter, parameters, eventsHandler);
					DocumentParameters currentDocument = converter.preprocessDocument(doc);
					ConversionResult result = converter.convert(currentDocument);

					// if result exist in expected
					if (expectedOutput.Exists) {
						// compare outputs
						//Assert
						string originalPluginResult = ReadFile(Path.Combine(conversionOutput, outputName));
						string currentResult = ReadFile(Path.Combine(expectedOutput.FullName, outputName));
						//
						Assert.AreEqual(originalPluginResult, currentResult, "Possible regression found");
					} else {
						// else move result folder to expected
						Directory.Move(conversionOutput, expectedOutput.FullName);
                    }
					
					//	compare output and expected
					
				} catch (Exception e) {
					testReport.WriteLine(e.Message);
				}

			}
			testReport.Close();


		}



		#region utilities

		public static List<string> getDocumentList(DirectoryInfo topDirectory) {
			List<string> result = new List<string>();
            foreach (string file in Directory.GetFiles(topDirectory.FullName,"*.docx")) {
				if(!Path.GetFileName(file).StartsWith("~")) {
					result.Add(file);
				}
				
			}
			foreach (string dir in Directory.GetDirectories(topDirectory.FullName)) {
				result.AddRange(getDocumentList(new DirectoryInfo(dir)));
            }
			return result;
		}

		protected string ReadFile(string filePath) {
			using (var stream = new StreamReader(File.Open(filePath, FileMode.Open))) {
				return stream.ReadToEnd();
			}
		}


		protected Application OpenMsWordDocument(string inputFile) {
			object falseValue = false;
			object trueValue = true;
			object missing = Type.Missing;

			CloseWord();

			_word = new Microsoft.Office.Interop.Word.Application();
			_word.Visible = false;
			_word.Activate();

			if (!string.IsNullOrEmpty(inputFile)) {
				object fileName = inputFile;
				_document = _word.Documents.Open(ref fileName, ref missing, ref falseValue, ref missing, ref missing, ref missing,
												 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
												 ref missing, ref missing, ref missing, ref missing);
				_document.Activate();
			}

			return _word;
		}


		protected void CloseWord() {
			object falseValue = false;
			object trueValue = true;
			object missing = Type.Missing;

			if (_word != null) {
				_word.Quit(ref falseValue, ref missing, ref missing);
				_word = null;
			}

			//Termenate word process if this was not close in previous step
			Process[] wordProcesses = Process.GetProcessesByName("WINWORD");
			foreach (var wordProcess in wordProcesses) {
				try {
					wordProcess.Kill();
				} catch (Exception ex) {
					Console.WriteLine("Can'not to terminate word process : " + ex.Message);
				}
			}
		}
		#endregion
	}
}
