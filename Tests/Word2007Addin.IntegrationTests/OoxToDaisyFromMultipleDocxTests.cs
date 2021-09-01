using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Daisy.SaveAsDAISY.Addins.Word2007;
using Extensibility;
using NUnit.Framework;
using Microsoft.Office.Interop.Word;
using Daisy.SaveAsDAISY.Conversion;
using System.Reflection;
using Daisy.SaveAsDAISY.Conversion.Events;

namespace Word2007Addin.IntegrationTests {
    [TestFixture]
    public class OoxToDaisyFromMultipleDocxTests : OoxToDaisyTestsBase {
        #region Overrides of OoxToDaisyTestsBase

        [TearDown]
        public override void TearDown() {
            base.TearDown();
        }

        [OneTimeTearDown]
        public override void FixtureTearDown() {
            base.FixtureTearDown();
        }

        [OneTimeSetUp]
        public override void FixtureSetUp() {
            base.FixtureSetUp();
        }

        #endregion

        [Test]
        public void Test1() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 1\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\1\Doc 1-2.xml").FullName;

            //NOTE: because translation contain validation errors output file will not created. And so we will check transitional output
            //string transitionalOuput = new FileInfo(@"output\1\F 2.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 1\Output\Doc 1-2.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();

            parameters
                .withParameter("Title", "m 1-2")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "1-2")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);


            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test1 failed.");
        }

        [Test]
        public void Test2() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 2\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\2\Doc 2-3.xml").FullName;

            //NOTE: because translation contain validation errors output file will not created. And so we will check transitional output
            //string transitionalOuput = new FileInfo(@"output\2\F 2.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 2\Output\Doc 2-3.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 2-3")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "2-3")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test2 failed.");
        }

        [Test]
        public void Test3() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 3\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\3\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 3\Output\Doc 3-4.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 3-4")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "3-4")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test3 failed.");
        }

        [Test]
        public void Test4() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 4\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\4\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 4\Output\Doc4-5.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 4-5")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "4-5")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test4 failed.");
        }

        [Test]
        public void Test5() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 5\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\5\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 5\Output\Doc 5-6.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 5-6")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "5-6")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test5 failed.");
        }

        [Test]
        public void Test6() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 6\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\6\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 6\Output\Doc 6-7.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 6-7")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "6-7")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test6 failed.");
        }

        [Test]
        public void Test7() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 7\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\7\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 7\Output\Doc 7-8.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 7-8")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "7-8")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test7 failed.");
        }

        [Test]
        public void Test8() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 8\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\8\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 8\Output\Doc 8-9.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 8-9")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "8-9")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test8 failed.");
        }

        [Test]
        public void Test9() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 9\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\9\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 9\Output\Doc 9-10.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 9-10")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "9-10")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test9 failed.");
        }

        [Test]
        public void Test10() {
            //Arrange
            DirectoryInfo inputDirectory = new DirectoryInfo(@"TestData\FromMultipleDocx\Test 10\input");
            ArrayList inputFiles = GetInputFilesList(inputDirectory);

            string outputFilePath = new FileInfo(@"output\10\Doc.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromMultipleDocx\Test 10\Output\Doc 10-1.xml").FullName;

            ConversionParameters parameters = new ConversionParameters();
            parameters
                .withParameter("Title", "m 10-1")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "10-1")
                .withParameter("Version", OfficeVersion)
                .withParameter("ParseSubDocuments", "Yes")
                .withParameter("Subject", "m");

            //Act
            convertAndMergeDocumentsToDaisy(inputFiles, outputFilePath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Multiple Docx Test10 failed.");
        }

        #region help methods

        private ArrayList GetInputFilesList(DirectoryInfo inputDirectory) {
            ArrayList inputFiles = new ArrayList();
            foreach (var inputFile in inputDirectory.GetFiles()) {
                if (!inputFile.Name.StartsWith("~")) {
                    inputFiles.Add(inputFile.FullName);
                }
            }
            return inputFiles;
        }

        public void convertAndMergeDocumentsToDaisy(ArrayList inputFiles, string ouputFilePath, ConversionParameters parameters) {
            Application word = OpenMsWordDocument(string.Empty);

            Connect connect = new Connect();

            Array array = new object[0];
            connect.OnConnection(word, ext_ConnectMode.ext_cm_External, null, ref array);

            SilentEventsHandler eventsHandler = new SilentEventsHandler();
            // TODO fix test
            /*connect.convertAndMergeDocumentsToDTBook(
                eventsHandler,
                inputFiles,
                ouputFilePath,
                string.Empty,
                "DaisyTabMultiple",
                parameters.ConversionParametersHash
            );*/
        }

        #endregion
    }
}