using System;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.Addins.Word2007 {
    public partial class ExceptionReport : Form {
        private Exception ExceptionRaised { get;  }
        public ExceptionReport(Exception raised) {
            InitializeComponent();
            this.ExceptionRaised = raised;
            this.ExceptionMessage.Text = raised.Message + "\r\nStacktrace:\r\n" + raised.StackTrace;
        }

        private void SendReport_Click(object sender, EventArgs evt) {
            StringBuilder message = new StringBuilder("The following exception was reported by a user using the saveAsDaisy addin:\r\n");
            message.AppendLine(this.ExceptionRaised.Message);
            message.AppendLine();
            message.Append(this.ExceptionRaised.StackTrace);
            Exception e = ExceptionRaised;
            while (e.InnerException != null) {
                e = e.InnerException;
                message.AppendLine(" - Inner exception : " + e.Message);
                message.AppendLine();
                message.Append(this.ExceptionRaised.StackTrace);
            }

            message.AppendLine();
            message.AppendLine("Addin version - " + typeof(ExceptionReport).Assembly.GetName().Version);
            
            string mailto = string.Format(
                "mailto:{0}?Subject={1}&Body={2}",
                "daisy-pipeline@mail.daisy.org",
                "Unhandled exception report in the SaveAsDAISY addin",
                message.ToString()
                );
            mailto = Uri.EscapeUriString(mailto);
            //MessageBox.Show(message.ToString());
            Process.Start(mailto);
        }
    }
}
