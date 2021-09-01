using System;
using System.IO;
using Daisy.SaveAsDAISY.Addins.Word2007;
using Extensibility;
using Microsoft.Office.Interop.Word;
using NUnit.Framework;
using Daisy.SaveAsDAISY.Conversion;

namespace Word2007Addin.IntegrationTests {
    [TestFixture]
    public class OoxToDaisyFromSingleDocxTests : OoxToDaisyTestsBase {
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

        /// <summary>
        /// Output should be equal to TestData/FromSingleDocx/Test1/Output1/F1.xml
        /// </summary>
        [Test]
        public void Test1() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 1\Input\F1.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\1").FullName;
            string outputFilePath = new FileInfo(@"output\1\F1.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 1\Output\F1.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Help")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "1")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test1 failed.");

        }

        [Test]
        public void Test2() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 2\Input\F 2.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\2").FullName;
            string outputFilePath = new FileInfo(@"output\2\F 2.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 2\Output\F 2.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Document")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "2")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test2 failed.");
        }

        [Test]
        public void Test3() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 3\Input\F 3.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\3").FullName;
            string outputFilePath = new FileInfo(@"output\3\F 3.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 3\Output\F 3.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Testing")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "3")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test3 failed.");
        }

        [Test]
        public void Test4() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 4\Input\F4.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\4").FullName;
            string outputFilePath = new FileInfo(@"output\4\F4.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 4\Output\F4.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Blood")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "4")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test4 failed.");
        }

        [Test]
        public void Test5() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 5\Input\F5.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\5").FullName;
            string outputFilePath = new FileInfo(@"output\5\F5.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 5\Output\F5.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Service")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "5")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test5 failed.");
        }

        [Test]
        public void Test6() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 6\Input\F6.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\6").FullName;
            string outputFilePath = new FileInfo(@"output\6\F6.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 6\Output\F6.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Belarus")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "6")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test6 failed.");
        }

        [Test]
        public void Test7() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 7\Input\F7.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\7").FullName;
            string outputFilePath = new FileInfo(@"output\7\F7.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 7\Output\F7.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Blog")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "7")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test7 failed.");
        }

        [Test]
        public void Test8() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 8\Input\F8.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\8").FullName;
            string outputFilePath = new FileInfo(@"output\8\F8.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 8\Output\F8.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Trial")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "8")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test8 failed.");
        }

        [Test]
        public void Test9() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 9\Input\F9.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\9").FullName;
            string outputFilePath = new FileInfo(@"output\9\F9.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 9\Output\F9.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Tool")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "9")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test9 failed.");
        }

        [Test]
        public void Test10() {
            //Arrange
            string inputFile = new FileInfo(@"TestData\FromSingleDocx\Test 10\Input\F10.docx").FullName;
            string outputDirectoryPath = new DirectoryInfo(@"output\10").FullName;
            string outputFilePath = new FileInfo(@"output\10\F10.xml").FullName;
            string originalOutputPath = new FileInfo(@"TestData\FromSingleDocx\Test 10\Output\F10.xml").FullName;


            ConversionParameters parameters = new ConversionParameters();
            parameters.withParameter("OutputFile", outputDirectoryPath)
                .withParameter("Title", "Problem")
                .withParameter("Creator", "Balandin Vyacheslav")
                .withParameter("Publisher", "Pruchkovskaya")
                .withParameter("UID", "10")
                .withParameter("TrackChangesFlag", "NoTrack")
                .withParameter("Version", OfficeVersion)
                .withParameter("MasterSubFlag", "No")
                .withParameter("Subject", string.Empty);

            //Act
            SaveAsSingleDaisy(inputFile, outputDirectoryPath, parameters);

            //Assert
            string originalPluginResult = ReadFile(originalOutputPath);
            string currentResult = ReadFile(outputFilePath);

            Assert.AreEqual(originalPluginResult, currentResult, "From Single Docx Test10 failed.");
        }

        #region help methods

        public void SaveAsSingleDaisy(string inputFile, string ouputDirectoryPath, ConversionParameters parameters) {
            Application word = OpenMsWordDocument(inputFile);

            Connect connect = new Connect();

            Array array = new object[0];
            connect.OnConnection(word, ext_ConnectMode.ext_cm_External, null, ref array);
            // TODO fix test
            /*connect.convertDocumentToDTBook(
                new SilentEventsHandler(),
                word.ActiveDocument,
                "DaisySingle",
                parameters,
                ouputDirectoryPath);*/

            
        }



        #endregion
    }
}