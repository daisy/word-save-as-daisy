using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour ExceptionReport.xaml
    /// </summary>
    public partial class ExceptionReport : Window
    {
        public ExceptionReport()
        {
            InitializeComponent();
        }

        private Exception ExceptionRaised { get; }
        public ExceptionReport(Exception raised)
        {
            string programFiles = System.Environment.ExpandEnvironmentVariables("%ProgramW6432%");
            var location = Assembly.GetCallingAssembly().Location;
            var isProgramFiles = location.StartsWith(programFiles);
            InitializeComponent();
            this.ExceptionRaised = raised;

            Exception current = raised;
            string exceptionText = current.Message;
            string trace = raised.StackTrace;
            while (current.InnerException != null) {
                current = current.InnerException;
                exceptionText += "\r\n- " + current.Message;
                trace += "\r\n" + current.StackTrace;
            }

            this.ExceptionMessage.Text = string.Format(
                "- Addin Version: {0}\r\n" +
                "- Running Architecture: {1}\r\n" +
                "- OS: {2}\r\n" +
                "- User or system wide install: {3}\r\n" +
                "- Addin settings:\r\n" +
                "```xml\r\n{6}\r\n```\r\n" +
                "{4}\r\n\r\n" +
                "Stacktrace:\r\n" +
                "```\r\n" +
                "{5}\r\n" +
                "```",
                typeof(ExceptionReport).Assembly.GetName().Version,
                System.Environment.Is64BitProcess ? "x64" : "x86",
                System.Runtime.InteropServices.RuntimeInformation.OSDescription,
                isProgramFiles ? "admin" : "user",
                exceptionText,
                trace,
                ConverterSettings.Instance.AsXML().Replace("\r\n", "").Replace("\t", "")
                );
        }

        private void SendReport_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder message = new StringBuilder("An exception was reported by the saveAsDaisy addin.\r\n");
            message.Append(ExceptionMessage.Text);

            string reportUrl = string.Format(
                "https://github.com/daisy/word-save-as-daisy/issues/new?title={0}&body={1}",
                Uri.EscapeDataString(ExceptionRaised.Message),
                Uri.EscapeDataString(message.ToString())
                );
            Process.Start(reportUrl);
        }

        private void Message_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://github.com/daisy/word-save-as-daisy/issues");
        }

        private void CheckForSimilarIssues_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://github.com/daisy/word-save-as-daisy/issues?q=is%3Aissue+" + Uri.EscapeDataString(ExceptionRaised.Message));
        }

    }
}
