using System;
using System.Diagnostics;
using System.IO;
using NUnit.Framework;
using Microsoft.Office.Interop.Word;

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


		[TestFixtureSetUp]
		public virtual void FixtureSetUp()
		{
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

		[TestFixtureTearDown]
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