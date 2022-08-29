using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.CommandLineTool {


    public class Report {

        enum Label {
            ERROR,
            WARNING,
            INFO,
            DEBUG
        };

        public const int DEBUG_LEVEL = 1;
        public const int INFO_LEVEL = 2;
        public const int WARNING_LEVEL = 3;
        public const int ERROR_LEVEL = 4;

        private StreamWriter writer = null;
        private int level = INFO_LEVEL;

        public Report(string filename, int level) {
            this.level = level;
            if (filename != null) {
                this.writer = new StreamWriter(new FileStream(filename, FileMode.Create, FileAccess.Write));
                Console.WriteLine("Using report file: " + filename);
            }
        }

        /*Function which adds comment*/
        public void AddComment(string message) {
            string text = "***" + message;

            if (this.writer != null) {
                this.writer.WriteLine(text);
                this.writer.Flush();
            }
            Console.WriteLine(text);
        }

        public void AddLog(string filename, string message, int level) {
            String text = "";
            if (level >= this.level) {
                string label = null;
                switch (level) {
                    case 4:
                        label = Label.ERROR.ToString();
                        break;
                    case 3:
                        label = Label.WARNING.ToString();
                        break;
                    case 2:
                        label = Label.INFO.ToString();
                        break;
                    default:
                        label = Label.DEBUG.ToString();
                        break;
                }
                if (filename != "")
                    text = "[" + label + "]" + "[" + filename + "] " + message;
                else
                    text = "[" + label + "]" + message;

                if (this.writer != null) {
                    this.writer.WriteLine(text);
                    this.writer.Flush();
                }
                Console.WriteLine(text);
            }
        }

        public void Close() {
            if (this.writer != null) {
                this.writer.Close();
                this.writer = null;
            }
        }

    }

}
