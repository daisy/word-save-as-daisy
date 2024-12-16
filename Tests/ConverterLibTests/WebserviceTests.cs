using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.IO;

namespace ConverterLibTests
{
    [TestClass]
    public class WebserviceTests
    {

        
        [TestMethod]
        public void AppIsInstalled()
        {
            Assert.IsTrue(PipelineApp.IsInstalled());
        }

        [TestMethod]
        public void AppIsRunning()
        {
            if(!PipelineApp.IsRunning())
            {
                PipelineApp.Start();
            }
            Assert.IsTrue(PipelineApp.IsRunning());
        }

        [TestMethod]
        public void EngineIsAlive()
        {
            Webservice ws = !PipelineApp.IsRunning() ? PipelineApp.Start() : PipelineApp.FindWebservice();
            Assert.IsNotNull(ws);
            Alive testAlive = ws.FetchAlive();
            Assert.IsNotNull(testAlive);
        }

        [TestMethod]
        public void GetScripts()
        {
            Webservice test = PipelineApp.FindWebservice();
            Assert.IsNotNull(test);
            ScriptData[] testScripts = test.FetchScripts();
            Assert.IsNotNull(testScripts);
        }

        [TestMethod]
        [DeploymentItem(@"test-files\\default_sample_for_dtbook_conversion.docx", "test-files")]
        public void ExecuteWordToDtbook()
        {
            Assert.IsTrue(File.Exists("test-files\\default_sample_for_dtbook_conversion.docx"));
            Webservice ws = !PipelineApp.IsRunning() ? PipelineApp.Start() : PipelineApp.FindWebservice();
            Assert.IsNotNull(ws);
            Assert.IsTrue(ws.FetchAlive().alive);
            ConsoleEventsHandler e = new ConsoleEventsHandler();
            WordToDtbook script = new WordToDtbook(e);
            script.Parameters["input"].ParameterValue = Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx");
            script.Parameters["output"].ParameterValue = Path.GetFullPath("test-results");
            if (!Directory.Exists(Path.GetFullPath("test-results"))) {
                Directory.CreateDirectory(Path.GetFullPath("test-results"));
            }
            script.TryExecuteOnAppWebservice(Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx"), false);
            Assert.IsTrue(File.Exists("test-results\\default_sample_for_dtbook_conversion.xml"));
        }

        [TestMethod]
        [DeploymentItem(@"test-files\\default_sample_for_dtbook_conversion.docx", "test-files")]
        public void ExecuteWordToCleanedDtbookChain()
        {
            Assert.IsTrue(File.Exists("test-files\\default_sample_for_dtbook_conversion.docx"));
            ConsoleEventsHandler e = new ConsoleEventsHandler();
            WordToCleanedDtbook script = new WordToCleanedDtbook(e);
            script.Parameters["input"].ParameterValue = Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx");
            script.Parameters["output"].ParameterValue = Path.GetFullPath("test-results");
            script.ExecuteScript(Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx"), false);

            // Search a file default_sample_for_dtbook_conversion.xml somewhere under the test-results folder
            DirectoryInfo results = new DirectoryInfo(Path.GetFullPath("test-results"));
            Assert.IsTrue(results.Exists);
            Assert.IsTrue(results.GetFiles("default_sample_for_dtbook_conversion.xml", SearchOption.AllDirectories).Any());
        }

        [TestMethod]
        [DeploymentItem(@"test-files\\default_sample_for_dtbook_conversion.docx", "test-files")]
        public void ExecuteWordToDaisy3Chain()
        {
            Assert.IsTrue(File.Exists("test-files\\default_sample_for_dtbook_conversion.docx"));
            ConsoleEventsHandler e = new ConsoleEventsHandler();
            WordToDaisy3 script = new WordToDaisy3(e);
            script.Parameters["input"].ParameterValue = Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx");
            script.Parameters["output"].ParameterValue = Path.GetFullPath("test-results");
            script.Parameters["audio"].ParameterValue = false;
            script.ExecuteScript(Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx"), false);

            // Search a file default_sample_for_dtbook_conversion.xml somewhere under the test-results folder
            DirectoryInfo results = new DirectoryInfo(Path.GetFullPath("test-results"));
            Assert.IsTrue(results.Exists);
            //Assert.IsTrue(results.GetFiles("default_sample_for_dtbook_conversion.xml", SearchOption.AllDirectories).Any());
        }

        [TestMethod]
        [DeploymentItem(@"test-files\\default_sample_for_dtbook_conversion.docx", "test-files")]
        public void ExecuteWordToEpub3Chain()
        {
            Assert.IsTrue(File.Exists("test-files\\default_sample_for_dtbook_conversion.docx"));
            ConsoleEventsHandler e = new ConsoleEventsHandler();
            WordToEpub3 script = new WordToEpub3(e);
            script.Parameters["input"].ParameterValue = Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx");
            script.Parameters["output"].ParameterValue = Path.GetFullPath("test-results");
            script.Parameters["audio"].ParameterValue = false;
            script.ExecuteScript(Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx"), false);

            // Search a file default_sample_for_dtbook_conversion.xml somewhere under the test-results folder
            DirectoryInfo results = new DirectoryInfo(Path.GetFullPath("test-results"));
            Assert.IsTrue(results.Exists);
            //Assert.IsTrue(results.GetFiles("default_sample_for_dtbook_conversion.xml", SearchOption.AllDirectories).Any());
        }

        [TestMethod]
        [DeploymentItem(@"test-files\\default_sample_for_dtbook_conversion.docx", "test-files")]
        public void ExecuteWordToDaisy202Chain()
        {
            Assert.IsTrue(File.Exists("test-files\\default_sample_for_dtbook_conversion.docx"));
            ConsoleEventsHandler e = new ConsoleEventsHandler();
            WordToDaisy202 script = new WordToDaisy202(e);
            script.Parameters["input"].ParameterValue = Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx");
            script.Parameters["output"].ParameterValue = Path.GetFullPath("test-results");
            script.Parameters["audio"].ParameterValue = false;
            script.ExecuteScript(Path.GetFullPath("test-files\\default_sample_for_dtbook_conversion.docx"), false);

            // Search a file default_sample_for_dtbook_conversion.xml somewhere under the test-results folder
            DirectoryInfo results = new DirectoryInfo(Path.GetFullPath("test-results"));
            Assert.IsTrue(results.Exists);
            //Assert.IsTrue(results.GetFiles("default_sample_for_dtbook_conversion.xml", SearchOption.AllDirectories).Any());
        }


        private static string GetCommandLine(Process process)
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + process.Id))
            using (ManagementObjectCollection objects = searcher.Get()) {
                return objects.Cast<ManagementBaseObject>().SingleOrDefault()?["CommandLine"]?.ToString();
            }

        }
    }
}
