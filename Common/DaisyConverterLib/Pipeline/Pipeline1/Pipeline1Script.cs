using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;


namespace Daisy.SaveAsDAISY.Conversion
{
    public class Pipeline1Script : Script
    {
        private string m_ScriptFilePath;
 
        public Pipeline1Script(string ScriptPath)
        {
            this._parameters = new Dictionary<string, ScriptParameter>();
            m_ScriptFilePath = ScriptPath;
            XmlTextReader reader = new XmlTextReader(m_ScriptFilePath);
            reader.XmlResolver = null;

            XmlDataDocument m_ScriptDocument = new XmlDataDocument();
            m_ScriptDocument.XmlResolver = null;
            m_ScriptDocument.Load(reader);
            reader.Close();

            //populate parameters list
            XmlNodeList CompleteNodeList = m_ScriptDocument.GetElementsByTagName("parameter");
            ScriptParameter p = null;
            foreach (XmlNode n in CompleteNodeList) {
                if (n.Attributes.Count > 0) {
                    p = new ScriptParameter(n);
                    this._parameters.Add(p.Name,p);
                }
            }
        }


       

        /// <summary>
        ///  executes script
        /// </summary>
        public override void ExecuteScript(string inputPath, bool isQuite)
        {
            string Param = "\"" + "input=" +  inputPath +"," ;
        	//string outputPath = string.Empty;

            int i = 0;
            foreach (var kv in Parameters)
            {
                ScriptParameter p = kv.Value;
                i++;
                if (p.IsParameterRequired && p.Name!="input")
                {
                    Param = Param + p.Name + "=" + p.ParameterValue + ",";
					if(p.Name == "output" || p.Name == "outputPath")
						output = p.ParameterValue;

                }
            }
            string str1 = Param;
            string str2 = str1.Substring(1,str1.LastIndexOf(",")-1);
            str2 = '"'+str2 + "\"";

            // invoke the script
            string PipelineFilePath = Path.Combine(Directory.GetParent(m_ScriptFilePath).Parent.FullName, "pipeline-lite.exe");
            Process PipelineProcess = new Process();
            bool displayOutputWindow = false;
            PipelineProcess.StartInfo.CreateNoWindow = !displayOutputWindow;
            PipelineProcess.StartInfo.UseShellExecute = displayOutputWindow;
            PipelineProcess.StartInfo.ErrorDialog = true;

            if(this.OnPipelineError != null) {
                // redirect pipeline error
                PipelineProcess.StartInfo.RedirectStandardError = true;
                PipelineProcess.ErrorDataReceived += (sender, args) => {
                    this.OnPipelineError(args.Data);
                };
                
            }

            if (this.OnPipelineOutput != null) {
                // Redirect pipeline output
                PipelineProcess.StartInfo.RedirectStandardOutput = true;
                PipelineProcess.OutputDataReceived += (sender, args) => {
                    this.OnPipelineOutput(args.Data);
                };
            }


            PipelineProcess.StartInfo.FileName = PipelineFilePath;
            PipelineProcess.StartInfo.Arguments = (isQuite ? "--quit " : string.Empty) + "--execute --script \"" + m_ScriptFilePath + "\" --params " + str2;
#if DEBUG
            if(this.OnPipelineOutput != null) {
                this.OnPipelineOutput("Launching " + PipelineProcess.StartInfo.FileName + " " + PipelineProcess.StartInfo.Arguments);
            }
#endif
            PipelineProcess.StartInfo.WorkingDirectory = Directory.GetParent(Directory.GetParent(m_ScriptFilePath).FullName).FullName;

            try
            {
                // Add pipeline logging
                PipelineProcess.Start();
                if (this.OnPipelineOutput != null) PipelineProcess.BeginOutputReadLine();
                if (this.OnPipelineError != null) PipelineProcess.BeginErrorReadLine();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
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

			if(PipelineProcess.ExitCode == 0 && !isQuite && !string.IsNullOrEmpty(output))
				System.Diagnostics.Process.Start(output);

        }

    }
}
