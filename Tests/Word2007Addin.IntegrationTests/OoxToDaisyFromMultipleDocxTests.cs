using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using DaisyWord2007AddIn;
using Extensibility;
using NUnit.Framework;
using Microsoft.Office.Interop.Word;
using Sonata.DaisyConverter.DaisyConverterLib.Converters;
using System.Reflection;

namespace Word2007Addin.IntegrationTests
{
	[TestFixture]
	public class OoxToDaisyFromMultipleDocxTests : OoxToDaisyTestsBase
	{
		#region Overrides of OoxToDaisyTestsBase

		[TearDown]
		public override void TearDown()
		{
			base.TearDown();
		}

		[OneTimeTearDown]
		public override void FixtureTearDown()
		{
			base.FixtureTearDown();
		}

		[OneTimeSetUp]
		public override void FixtureSetUp()
		{
			base.FixtureSetUp();
		}

		#endregion

		[Test]
		public void Test1()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 1\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\1\Doc 1-2.xml").FullName;

			//NOTE: because translation contain validation errors output file will not created. And so we will check transitional output
			//string transitionalOuput = new FileInfo(@"output\1\F 2.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 1\Output\Doc 1-2.xml").FullName;


			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 1-2")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("1-2")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);


			//Assert
			string originalPluginResult = ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test1 failed.");
		}

		[Test]
		public void Test2()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 2\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\2\Doc 2-3.xml").FullName;

			//NOTE: because translation contain validation errors output file will not created. And so we will check transitional output
			//string transitionalOuput = new FileInfo(@"output\2\F 2.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 2\Output\Doc 2-3.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 2-3")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("2-3")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test2 failed.");
		}

		[Test]
		public void Test3()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 3\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\3\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 3\Output\Doc 3-4.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 3-4")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("3-4")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test3 failed.");
		}

		[Test]
		public void Test4()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 4\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\4\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 4\Output\Doc4-5.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 4-5")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("4-5")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test4 failed.");
		}

		[Test]
		public void Test5()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 5\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\5\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 5\Output\Doc 5-6.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 5-6")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("5-6")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test5 failed.");
		}

		[Test]
		public void Test6()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 6\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\6\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 6\Output\Doc 6-7.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 6-7")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("6-7")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test6 failed.");
		}

		[Test]
		public void Test7()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 7\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\7\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 7\Output\Doc 7-8.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 7-8")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("7-8")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test7 failed.");
		}

		[Test]
		public void Test8()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 8\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\8\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 8\Output\Doc 8-9.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 8-9")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("8-9")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test8 failed.");
		}

		[Test]
		public void Test9()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 9\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\9\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 9\Output\Doc 9-10.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 9-10")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("9-10")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test9 failed.");
		}

		[Test]
		public void Test10()
		{
			//Arrange
			DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 10\input");
			ArrayList inputFiles = GetInputFilesList(inputDirectory);

			string outputFilePath = new FileInfo(@"output\10\Doc.xml").FullName;
			string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 10\Output\Doc 10-1.xml").FullName;

			TranslationParametersBuilder preporator = new TranslationParametersBuilder();
			preporator
				.WithTitle("m 10-1")
				.WithCreator("Balandin Vyacheslav")
				.WithPublisher("Pruchkovskaya")
				.WithUID("10-1")
				.WithVersion(OfficeVersion)
				.WithMasterSubFlag("Yes")
				.WithSubject("m");

			//Act
			SaveAsMultipleDaisy(inputFiles, outputFilePath, preporator);

			//Assert
			string originalPluginResult =ReadFile(originalOutputPath);
			string currentResult = ReadFile(outputFilePath);

			Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test10 failed.");
		}

		#region help methods

		private ArrayList GetInputFilesList(DirectoryInfo inputDirectory)
		{
			ArrayList inputFiles = new ArrayList();
			foreach (var inputFile in inputDirectory.GetFiles())
			{
                if (!inputFile.Name.StartsWith("~")) {
					inputFiles.Add(inputFile.FullName);
				}
			}
			return inputFiles;
		}

		public void SaveAsMultipleDaisy(ArrayList inputFiles, string ouputFilePath, TranslationParametersBuilder preporator)
		{
			Application word = OpenMsWordDocument(string.Empty);

			Connect connect = new Connect();

			Array array = new object[0];
			connect.OnConnection(word, ext_ConnectMode.ext_cm_External, null, ref array);

			connect.SignleDaisyFromMultipleInQuiteMode(inputFiles, ouputFilePath, preporator);
		}

		#endregion
	}
}