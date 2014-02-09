/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Reflection;
using System.Collections;
using System.Xml.Schema;
using System.Windows.Forms;


namespace Sonata.DaisyConverter.DaisyConverterLib
{
    /// <summary>
    /// Core conversion methods 
    /// </summary>
    public abstract class AbstractConverter
    {
        private const string SOURCE_XML = "source.xml";
        private bool isDirectTransform = true;
        private ArrayList skipedPostProcessors = null;
        private string externalResource = null;
        private Assembly resourcesAssembly;
        private Hashtable compiledProcessors;
        private static bool isValid;
        private static string error;
        ArrayList MathList8879, MathList9573, MathListmathml;
        private ArrayList fidilityLoss = new ArrayList();
        String errorText = "";

        protected AbstractConverter(Assembly resourcesAssembly)
        {
            this.resourcesAssembly = resourcesAssembly;
            this.skipedPostProcessors = new ArrayList();
            this.compiledProcessors = new Hashtable();
            AddMathmlDtds();
        }

        public bool DirectTransform
        {
            set { this.isDirectTransform = value; }
            get { return this.isDirectTransform; }
        }

        public ArrayList SkipedPostProcessors
        {
            set { this.skipedPostProcessors = value; }
        }

        public string ExternalResources
        {
            set { this.externalResource = value; }
            get { return this.externalResource; }
        }

        public ArrayList FidilityLoss
        {
            get
            {
                return fidilityLoss;
            }
        }

        /// <summary>
        /// Pull the chain of post processors for the direct conversion
        /// </summary>
        protected virtual string[] DirectPostProcessorsChain
        {
            get { return null; }
        }

        /// <summary>
        /// Pull the chain of post processors for the reverse conversion
        /// </summary>
        protected virtual string[] ReversePostProcessorsChain
        {
            get { return null; }
        }

        /// <summary>
        /// Pull an XmlUrlResolver for embedded resources
        /// </summary>
        private XmlUrlResolver ResourceResolver
        {
            get
            {
                if (this.ExternalResources == null)
                {
                    return new EmbeddedResourceResolver(this.resourcesAssembly,
                      this.GetType().Namespace, this.DirectTransform);
                }
                else
                {
                    return new SharedXmlUrlResolver(this.DirectTransform);
                }
            }
        }

        /// <summary>
        /// Pull the input xml document to the xsl transformation
        /// </summary>
        private XmlReader Source
        {
            get
            {
                XmlReaderSettings xrs = new XmlReaderSettings();
                // do not look for DTD
                xrs.ProhibitDtd = true;
                if (this.ExternalResources == null)
                {
                    xrs.XmlResolver = this.ResourceResolver;
                    return XmlReader.Create(SOURCE_XML, xrs);
                }
                else
                {
                    return XmlReader.Create(this.ExternalResources + "/" + SOURCE_XML, xrs);
                }
            }
        }


        private XslCompiledTransform Load(bool computeSize)
        {
            try
            {
                string xslLocation = "oox2Daisy.xsl";
                XPathDocument xslDoc = null;
                XmlUrlResolver resolver = this.ResourceResolver;

                if (this.ExternalResources == null)
                {
                    if (computeSize)
                    {
                        xslLocation = "oox2Daisy.xsl";
                    }
                    EmbeddedResourceResolver emr = (EmbeddedResourceResolver)resolver;
                    emr.IsDirectTransform = this.DirectTransform;
                    xslDoc = new XPathDocument(emr.GetInnerStream(xslLocation));
                }
                else
                {
                    xslDoc = new XPathDocument(this.ExternalResources + "/" + xslLocation);
                }

                if (!this.compiledProcessors.Contains(xslLocation))
                {

#if DEBUG
                    XslCompiledTransform xslt = new XslCompiledTransform(true);
#else
                    XslCompiledTransform xslt = new XslCompiledTransform();
#endif

                    // compile the stylesheet. 
                    // Input stylesheet, xslt settings and uri resolver are retrieve from the implementation class.
                    xslt.Load(xslDoc, this.XsltProcSettings, this.ResourceResolver);
                    this.compiledProcessors.Add(xslLocation, xslt);
                }
                return (XslCompiledTransform)this.compiledProcessors[xslLocation];
            }
            catch (Exception e)
            {
                string s;
                s = e.Message;
                return null;
            }

        }

        /// <summary>
        /// Pull the xslt settings
        /// </summary>
        private XsltSettings XsltProcSettings
        {
            get
            {
                // Enable xslt 'document()' function
                return new XsltSettings(true, false);
            }
        }

        public void ComputeSize(string inputFile, Hashtable table)
        {
            Transform(inputFile, null, table, null, true, "");
        }

        public void Transform(string inputFile, string outputFile, Hashtable table, Hashtable listMathMl, bool modeValue, string output_Pipeline)
        {
            fidilityLoss = new ArrayList();
            string tempInputFile = Path.GetTempFileName();
            string tempOutputFile = outputFile == null ? null : Path.GetTempFileName();
            try
            {
                File.Copy(inputFile, tempInputFile, true);
                File.SetAttributes(tempInputFile, FileAttributes.Normal);
                _Transform(inputFile, tempInputFile, tempOutputFile, outputFile, listMathMl, table, output_Pipeline);

                if (outputFile != null)
                {
                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }
                    File.Move(tempOutputFile, outputFile);

                    CopyDTDToDestinationfolder(outputFile);
                    CopyCSSToDestinationfolder(outputFile);
                    CopyMATHToDestinationfolder(outputFile);

                    if (modeValue)
                    {
                        XmlValidation(outputFile);

                    }

                    Int16 value = (Int16)outputFile.LastIndexOf("\\");
                    String tempStr = outputFile.Substring(0, value);
                    DeleteDTD(tempStr + "\\" + "dtbook-2005-3.dtd", outputFile, modeValue);
                    DeleteMath(tempStr, modeValue);
                }
            }
            finally
            {
                if (File.Exists(tempInputFile))
                {
                    try
                    {
                        File.Delete(tempInputFile);
                    }
                    catch (IOException)
                    {
                        Debug.Write("could not delete temporary input file");
                    }
                }
            }
        }


        public void DeleteDTD(String fileDTD, String fileName, bool value)
        {

            /*temperory solution - needs to be changed*/
            StreamReader reader = new StreamReader(fileName);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            if (value)
            {
                data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'");
                data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
                if (!data.Contains("</mml:math>"))
                {
                    data = data.Remove(203, 917);
                    data = data.Replace(Environment.NewLine + "<dtbook", "<dtbook");
                    data = data.Replace("xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"", "");
                }
                data = data.Replace("<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>", "<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">");
            }
            writer.Write(data);
            writer.Close();

            if (value)
            {
                if (File.Exists(fileDTD))
                {
                    File.Delete(fileDTD);
                }
            }
        }

        public void DeleteMath(String fileName, bool value)
        {
            if (value)
            {
                DeleteFile(fileName + "\\mathml2.DTD");
                DeleteFile(fileName + "\\mathml2-qname-1.mod");
                Directory.Delete(fileName + "\\iso8879", true);
                Directory.Delete(fileName + "\\iso9573-13", true);
                Directory.Delete(fileName + "\\mathml", true);
            }
        }

        public void DeleteFile(String file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        private void _Transform(String inputName, string inputFile, string outputFile, string actualoutputfile, Hashtable listMathMl, Hashtable table, string output_Pipeline)
        {
            // this throws an exception in the the following cases:
            // - input file is not a valid file
            // - input file is an encrypted file

            XmlReader source = null;
            XmlWriter writer = null;
            ZipResolver zipResolver = null;

            try
            {
                XslCompiledTransform xslt = this.Load(outputFile == null);
                zipResolver = new ZipResolver(inputFile);

                XsltArgumentList parameters = new XsltArgumentList();
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBackMaster);

                if (outputFile != null)
                {
                    if (actualoutputfile.ToLower().EndsWith(".xml"))
                    {
                        int length = actualoutputfile.LastIndexOf("\\");
                        string s1 = actualoutputfile.Substring(0, length);


                        DaisyClass val = new DaisyClass(inputName, inputFile, s1, listMathMl, zipResolver.Archive, output_Pipeline);
                        parameters.AddExtensionObject("urn:Daisy", val);


                        parameters.AddParam("outputFile", "", outputFile);

                        foreach (DictionaryEntry myEntry in table)
                        {
                            parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
                        }


                        XmlWriter finalWriter;
#if DEBUG
						StreamWriter streamWriter = new StreamWriter(outputFile, true, System.Text.Encoding.UTF8) {AutoFlush = true};
                    	finalWriter = new XmlTextWriter(streamWriter);
#else
                    	finalWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);
#endif

#if DEBUG
                    	Debug.WriteLine("OUTPUT FILE : '" + outputFile + "'");
#endif

                        writer = GetWriter(finalWriter);
                    }
                    else
                    {
                        DaisyClass val = new DaisyClass(inputName, inputFile, actualoutputfile, listMathMl, zipResolver.Archive, output_Pipeline);
                        parameters.AddExtensionObject("urn:Daisy", val);


                        parameters.AddParam("outputFile", "", outputFile);

                        foreach (DictionaryEntry myEntry in table)
                        {
                            parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
                        }


                        XmlWriter finalWriter;
                        finalWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);

                        writer = GetWriter(finalWriter);
                    }

                }
                else
                {

                    int length = inputFile.LastIndexOf("\\");
                    string s1 = inputFile.Substring(0, length);

                    DaisyClass val = new DaisyClass(inputName, inputFile, s1, listMathMl, zipResolver.Archive, output_Pipeline);
                    parameters.AddExtensionObject("urn:Daisy", val);

					if (table != null)
					{
						foreach (DictionaryEntry myEntry in table)
							parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
					}


                	writer = new XmlTextWriter(new StringWriter());
                }
                source = this.Source;
                // Apply the transformation

                xslt.Transform(source, parameters, writer, zipResolver);                
            }
            finally
            {
                if (writer != null)
                    writer.Close();
                if (source != null)
                    source.Close();
                if (zipResolver != null)
                    zipResolver.Dispose();
            }
        }

        private void MessageCallBack(object sender, XsltMessageEncounteredEventArgs e)
        {
            if (e.Message.StartsWith("progress:"))
            {
                if (progressMessageIntercepted != null)
                {
                    progressMessageIntercepted(this, null);
                }
            }
            else if (e.Message.StartsWith("translation.oox2Daisy."))
            {
                fidilityLoss.Add(e.Message);

                if (feedbackMessageIntercepted != null)
                {
                    feedbackMessageIntercepted(this, new DaisyEventArgs(e.Message));

                }
            }
        }

        private void MessageCallBackMaster(object sender, XsltMessageEncounteredEventArgs e)
        {
            if (e.Message.StartsWith("progress:"))
            {
                if (progressMessageInterceptedMaster != null)
                {
                    progressMessageInterceptedMaster(this, null);
                }
            }
        }

        private XmlWriter GetWriter(XmlWriter writer)
        {
            string[] postProcessors = this.DirectPostProcessorsChain;
            if (!this.isDirectTransform)
            {
                postProcessors = this.ReversePostProcessorsChain;
            }
            return InstanciatePostProcessors(postProcessors, writer);
        }


        private XmlWriter InstanciatePostProcessors(string[] procNames, XmlWriter lastProcessor)
        {
            XmlWriter currentProc = lastProcessor;
            if (procNames != null)
            {
                for (int i = procNames.Length - 1; i >= 0; --i)
                {
                    if (!Contains(procNames[i], this.skipedPostProcessors))
                    {
                        Type type = Type.GetType(procNames[i]);
                        object[] parameters = { currentProc };
                        XmlWriter newProc = (XmlWriter)Activator.CreateInstance(type, parameters);
                        currentProc = newProc;
                    }
                }
            }
            return currentProc;
        }

        private bool Contains(string processorFullName, ArrayList names)
        {
            foreach (string name in names)
            {
                if (processorFullName.Contains(name))
                {
                    return true;
                }
            }
            return false;
        }

        public delegate void MessageListener(object sender, EventArgs e);
        private event MessageListener progressMessageIntercepted;
        private event MessageListener progressMessageInterceptedMaster;
        private event MessageListener feedbackMessageIntercepted;
        private event MessageListener feedbackValidationIntercepted;

        public void AddProgressMessageListener(MessageListener listener)
        {
            progressMessageIntercepted += listener;
        }

        public void AddProgressMessageListenerMaster(MessageListener listener)
        {
            progressMessageInterceptedMaster += listener;
        }

        public void AddFeedbackMessageListener(MessageListener listener)
        {
            feedbackMessageIntercepted += listener;
        }

        public void AddFeedbackValidationListener(MessageListener listener)
        {
            feedbackValidationIntercepted += listener;
        }

        public void RemoveMessageListeners()
        {
            progressMessageIntercepted = null;
            feedbackMessageIntercepted = null;
            feedbackValidationIntercepted = null;
            progressMessageInterceptedMaster = null;
        }

        /* Function to Copy DTD file to the Output folder*/
        public void CopyDTDToDestinationfolder(String outputFile)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            string fileName = Path.GetDirectoryName(outputFile) + "\\dtbook-2005-3.dtd";
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("dtbook-2005-3.dtd"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            string data = reader.ReadToEnd();
            reader.Close();

            StreamWriter writer = new StreamWriter(fileName);
            writer.Write(data);
            writer.Close();

        }

        public void CopyCSSToDestinationfolder(String outputFile)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;
            string fileName = Path.GetDirectoryName(outputFile) + "\\dtbookbasic.css";
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("dtbookbasic.css"))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }
            StreamReader reader = new StreamReader(stream);
            StreamWriter writer = new StreamWriter(fileName);
            string data = reader.ReadToEnd();
            writer.Write(data);
            reader.Close();
            writer.Close();
        }

        public void CopyMATHToDestinationfolder(String outputFile)
        {
            string fileName = "";
            fileName = Path.GetDirectoryName(outputFile) + "\\mathml2-qname-1.mod";
            CopyingAssemblyFile(fileName, "mathml2-qname-1.mod");
            fileName = Path.GetDirectoryName(outputFile) + "\\mathml2.DTD";
            CopyingAssemblyFile(fileName, "mathml2.DTD");

            for (int i = 0; i < MathList8879.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso8879");
                fileName = Path.GetDirectoryName(outputFile) + "\\iso8879\\" + MathList8879[i].ToString();
                CopyingAssemblyFile(fileName, MathList8879[i].ToString());

            }

            for (int i = 0; i < MathList9573.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso9573-13");
                fileName = Path.GetDirectoryName(outputFile) + "\\iso9573-13\\" + MathList9573[i].ToString();
                CopyingAssemblyFile(fileName, MathList9573[i].ToString());
            }

            for (int i = 0; i < MathListmathml.Count; i++)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\mathml");
                fileName = Path.GetDirectoryName(outputFile) + "\\mathml\\" + MathListmathml[i].ToString();
                CopyingAssemblyFile(fileName, MathListmathml[i].ToString());

            }
        }

        public void CopyingAssemblyFile(String fileName, String indFilename)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            Stream stream = null;

            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(indFilename))
                {
                    stream = asm.GetManifestResourceStream(name);
                    break;
                }

            }

            StreamReader reader = new StreamReader(stream);
            StreamWriter writer = new StreamWriter(fileName);
            string data = reader.ReadToEnd();
            writer.Write(data);
            reader.Close();
            writer.Close();
        }

        /*Function to do validation of Output XML file with DTD*/
        public void XmlValidation(String outFile)
        {
            isValid = true;
            error = "";
            XmlTextReader xml = new XmlTextReader(outFile);
            XmlValidatingReader xsd = new XmlValidatingReader(xml);

            try
            {

                xsd.ValidationType = ValidationType.DTD;
                ArrayList errTxt = new ArrayList();
                for (int i = 0; i <= 4; i++)
                    errTxt.Add("");
                xsd.ValidationEventHandler += new ValidationEventHandler(MyValidationEventHandler);
                while (xsd.Read())
                {
                    errTxt[4] = errTxt[3];
                    errTxt[3] = errTxt[2];
                    errTxt[2] = errTxt[1];
                    errTxt[1] = errTxt[0];
                    errTxt[0] = xsd.ReadString();
                    errorText = "";
                    for (int i = 4; i >= 0; i--)
                        errorText = errorText + errTxt[i].ToString() + " ";
                    if (errorText.Contains("\n"))
                        errorText = errorText.Replace("\n", "");
                    if (errorText.Contains("\r"))
                        errorText = errorText.Replace("\r", "");
                    if (errorText.Length > 100)
                        errorText = errorText.Substring(0, 100);
                }
                xsd.Close();

                Stream stream = null;
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (name.EndsWith("Shematron.xsl"))
                    {
                        stream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }

                XmlReader rdr = XmlReader.Create(stream);
                XPathDocument doc = new XPathDocument(outFile);

                XslCompiledTransform trans = new XslCompiledTransform(true);
                trans.Load(rdr);

                XmlTextWriter myWriter = new XmlTextWriter(Path.GetDirectoryName(outFile) + "\\report.txt", null);
                trans.Transform(doc, null, myWriter);

                myWriter.Close();
                rdr.Close();

                StreamReader reader = new StreamReader(Path.GetDirectoryName(outFile) + "\\report.txt");
                if (!reader.EndOfStream)
                {
                    error += reader.ReadToEnd();

                    if (feedbackValidationIntercepted != null)
                    {
                        feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                    }
                }
                reader.Close();

                if (File.Exists(Path.GetDirectoryName(outFile) + "\\report.txt"))
                {
                    File.Delete(Path.GetDirectoryName(outFile) + "\\report.txt");
                }

                // Check whether the document is valid or invalid.
                if (isValid == false)
                {
                    if (feedbackValidationIntercepted != null)
                    {
                        feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                    }
                }
            }
            catch (UnauthorizedAccessException a)
            {
                xsd.Close();
                //dont have access permission
                error = a.Message;

                if (feedbackValidationIntercepted != null)
                {
                    feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                }
            }
            catch (Exception a)
            {
                xsd.Close();
                //and other things that could go wrong
                error = a.Message;

                if (feedbackValidationIntercepted != null)
                {
                    feedbackValidationIntercepted(this, new DaisyEventArgs(error));
                }
            }
        }


        /* Function to Capture all the Validity Errors*/
        public void MyValidationEventHandler(object sender, ValidationEventArgs args)
        {
            isValid = false;
            error += " Line Number : " + args.Exception.LineNumber + " and " +
             " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
             " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;
        }

        public void AddMathmlDtds()
        {
            MathList8879 = new ArrayList();
            MathList9573 = new ArrayList();
            MathListmathml = new ArrayList();

            MathList8879.Add("isobox.ent");
            MathList8879.Add("isocyr1.ent");
            MathList8879.Add("isocyr2.ent");
            MathList8879.Add("isodia.ent");
            MathList8879.Add("isolat1.ent");
            MathList8879.Add("isolat2.ent");
            MathList8879.Add("isonum.ent");

            MathList8879.Add("isopub.ent");

            MathListmathml.Add("mmlalias.ent");
            MathListmathml.Add("mmlextra.ent");

            MathList9573.Add("isoamsa.ent");
            MathList9573.Add("isoamsb.ent");
            MathList9573.Add("isoamsc.ent");
            MathList9573.Add("isoamsn.ent");
            MathList9573.Add("isoamso.ent");
            MathList9573.Add("isoamsr.ent");
            MathList9573.Add("isogrk3.ent");
            MathList9573.Add("isomfrk.ent");
            MathList9573.Add("isomopf.ent");
            MathList9573.Add("isomscr.ent");
            MathList9573.Add("isotech.ent");
        }

    }
}
