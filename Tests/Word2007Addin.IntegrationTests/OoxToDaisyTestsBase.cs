using System;
using System.Diagnostics;
using System.IO;
using NUnit.Framework;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Word2007Addin.IntegrationTests
{
	public abstract class OoxToDaisyTestsBase
	{
		protected Application _word;
		protected Document _document;
		public const string OfficeVersion = "14";

		protected void CloseWord()
		{
			object falseValue = false;
			object trueValue = true;
			object missing = Type.Missing;

			if (_word != null)
			{
				_word.Quit(ref falseValue, ref missing, ref missing);
				_word = null;
			}

			//Termenate word process if this was not close in previous step
			Process[] wordProcesses = Process.GetProcessesByName("WINWORD");
			foreach (var wordProcess in wordProcesses)
			{
				try
				{
					wordProcess.Kill();
				}
				catch (Exception ex)
				{
					Console.WriteLine("Can'not to terminate word process : " + ex.Message);
				}
			}
		}


		[OneTimeSetUp]
		public virtual void FixtureSetUp()
		{
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

			try
			{
				if (Directory.Exists("output"))
					Directory.Delete("output", true);

				Directory.CreateDirectory("output");
				for (int i = 1; i < 11; i++)
				{
					Directory.CreateDirectory(string.Format("output\\{0}", i));
				}
			}catch(Exception ex)
			{
				Console.WriteLine("Some error during output folders preparation : " + ex.Message);
			}
		}

		[OneTimeTearDown]
		public virtual void FixtureTearDown()
		{
			if (Directory.Exists("output"))
			{
				try
				{
					Directory.Delete("output", true);
				}
				catch (Exception ex)
				{
					Console.WriteLine("Can'not delete output folder : " + ex.Message);
				}
				
			}
			CloseWord();
		}

		[TearDown]
		public virtual void TearDown()
		{
			CloseWord();
		}

		protected string ReadFile(string filePath)
		{
			using (var stream = new StreamReader(File.Open(filePath, FileMode.Open)))
			{
				return stream.ReadToEnd();
			}
		}


		protected Application OpenMsWordDocument(string inputFile)
		{
			object falseValue = false;
			object trueValue = true;
			object missing = Type.Missing;

			CloseWord();

			_word = new Microsoft.Office.Interop.Word.Application();
			_word.Visible = true;
			_word.Activate();

			if (!string.IsNullOrEmpty(inputFile))
			{
				object fileName = inputFile;
				_document = _word.Documents.Open(ref fileName, ref missing, ref falseValue, ref missing, ref missing, ref missing,
												 ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
												 ref missing, ref missing, ref missing, ref missing);
				_document.Activate();
			}

			return _word;
		}
	}
}