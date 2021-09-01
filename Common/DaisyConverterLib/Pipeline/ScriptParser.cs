using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;


namespace Daisy.SaveAsDAISY.Conversion
{
    public class ScriptParser
    {
        private XmlDataDocument m_ScriptDocument;
        private string m_ScriptFilePath;
        private List<ScriptParameter> m_ParameterList;

        private string m_NiceName = "";

        /// <summary>
        /// List of parameters available in script
        /// </summary>
        public List<ScriptParameter> ParameterList {
            get { return m_ParameterList; }
        }

        public string NiceName { 
            get {
                if (m_NiceName == "") {
                    // I assume the taskscript to start by a nicename tag
                    XmlNode taskscript = m_ScriptDocument.GetElementsByTagName("taskScript")[0];
                    m_NiceName = taskscript.FirstChild.InnerText;
                }
                return m_NiceName;
            }
        }


        public ScriptParser(string ScriptPath)
        {
            m_ParameterList = new List<ScriptParameter>();
            m_ScriptFilePath = ScriptPath;
            XmlTextReader reader = new XmlTextReader(m_ScriptFilePath);
            reader.XmlResolver = null;

            m_ScriptDocument = new XmlDataDocument();
            m_ScriptDocument.XmlResolver = null;
            m_ScriptDocument.Load(reader);
            reader.Close();

            //populate parameters list
            PopulateParameterList();
        }



        /// <summary>
        /// <summary
        ///  populate parameter list
        /// </summary>
        private void PopulateParameterList()
        {
            XmlNodeList CompleteNodeList = m_ScriptDocument.GetElementsByTagName("parameter");
            ScriptParameter p = null;
            foreach (XmlNode n in CompleteNodeList)
            {
                if (n.Attributes.Count > 0)
                {
                    p = new ScriptParameter(n);
                    m_ParameterList.Add(p);
                }
            }

        }


        

        /// <summary>
        /// Executes script normal.
        /// </summary>
        /// <param name="inputPath"></param>
        public void ExecuteScript(string inputPath)
        {
            ExecuteScript(inputPath, false);
        }

        /// <summary>
        ///  executes script
        /// </summary>
        public void ExecuteScript(string inputPath, bool isQuite)
        {
            /*
            if (!File.Exists(m_InputFile.ParameterValue)
                            || !Directory.Exists(m_OutputDirectory.ParameterValue) || Directory.GetFiles(m_OutputDirectory.ParameterValue).Length > 0
                            || (m_BitRate.ParameterValue != "32" && m_BitRate.ParameterValue != "48" m_BitRate.ParameterValue != "64" && m_BitRate.ParameterValue != "128"))
            {
                throw new System.Exception( Localizer.Message ("Invalid_ScriptParameters")) ;
                return;
            }
             */



            //pipeline-lite.exe --execute  --script "scripts/Narrator-Autumn2008.taskScript" --params "input=1.xml,outputPath=C:\PipelineLite\pipeline-lite-ms-20080901\pipeline-lite-ms\output"

            string Param = "\"" + "input=" +  inputPath +"," ;
        	string outputPath = string.Empty;

            int i = 0;
            foreach (ScriptParameter p in ParameterList)
            {
                i++;
                if (p.IsParameterRequired && p.Name!="input")
                {
                    //Param = Param + " --\"" + p.ParameterName + "=" + p.ParameterValue + "\"";
                    //if(i!=ParameterList.Count-1)
                        Param = Param + p.Name + "=" + p.ParameterValue + ",";
                    //else
                      //  Param = Param + p.Name + "=" + p.ParameterValue + "\"";
					if(p.Name == "output" || p.Name == "outputPath")
						outputPath = p.ParameterValue;

                }
            }
            string str1 = Param;
            string str2 = str1.Substring(1,str1.LastIndexOf(",")-1);
            str2 = '"'+str2 + "\"";


            //foreach (ScriptParameter p in ParameterList)
            //{
            //    if (p.IsParameterRequired)
            //    {
            //        //Param = Param + " --\"" + p.ParameterName + "=" + p.ParameterValue + "\"";
            //        Param = Param + " \"" + p.Name + "=" + p.ParameterValue + "\"";
            //    }
            //}
            // invoke the script
            string PipelineFilePath = Path.Combine(Directory.GetParent(m_ScriptFilePath).Parent.FullName, "pipeline-lite.exe");
            Process PipelineProcess = new Process();
            bool displayOutputWindow = false;
            PipelineProcess.StartInfo.CreateNoWindow = !displayOutputWindow;
            PipelineProcess.StartInfo.UseShellExecute = displayOutputWindow;
            PipelineProcess.StartInfo.ErrorDialog = true;
            

            PipelineProcess.StartInfo.FileName = PipelineFilePath;
            PipelineProcess.StartInfo.Arguments = (isQuite ? "--quit " : string.Empty) + "--execute --script \"" + m_ScriptFilePath + "\" --params " + str2;
            PipelineProcess.StartInfo.WorkingDirectory = Directory.GetParent(Directory.GetParent(m_ScriptFilePath).FullName).FullName;

            try
            {
                PipelineProcess.Start();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            // TODO add pipeline output and error redirection to a log file for pipeline debugging ...
            PipelineProcess.WaitForExit();
            //Deleting the files
            if (File.Exists(inputPath))
            {
                File.Delete(inputPath);
            }
            int s = inputPath.LastIndexOf('\\');
            string strPath = inputPath.Substring(0, s);
            string[] files = Directory.GetFiles(strPath);
            foreach (string file in files)
            {
                if (file.Contains(".jpg") || file.Contains(".JPG") || file.Contains(".PNG") || file.Contains(".png"))
                {
                    File.Delete(file);
                }
                if (file.Contains("dtbookbasic.css"))
                {
                    File.Delete(file);
                }
            }

			if(PipelineProcess.ExitCode == 0 && !isQuite && !string.IsNullOrEmpty(outputPath))
				System.Diagnostics.Process.Start(outputPath);

        }

    }
}
