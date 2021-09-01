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
using System.IO.Packaging;
using System.Text;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Daisy.SaveAsDAISY.Conversion
{
	public class ValidationError {
		public readonly ValidationEventArgs error;
		public readonly string filePath;
		public readonly string referenceText;
		public ValidationError(ValidationEventArgs error, string filePath, string referenceText) {
			this.error = error;
			this.filePath = filePath;
			this.referenceText = referenceText;
		}

		public override string ToString() {
			return "File " + filePath + Environment.NewLine + 
				" Line Number : " + error.Exception.LineNumber + " - " + " Line Position : " + error.Exception.LinePosition + Environment.NewLine +
				" Message : " + error.Message + Environment.NewLine + 
				" Reference Text :  " + referenceText + Environment.NewLine;
		}

	}

	/// <summary>
	/// Core conversion class that convert a word file (on disk) to XML
	/// </summary>
	public class WordToDTBookXMLTransform
	{
		private const string SOURCE_XML = "source.xml";

		/// <summary>
		/// alldegedly allow to load
		/// </summary>
		private bool isDirectTransform = true;

		private ArrayList skipedPostProcessors = null;
		
		private string externalResource = null;
		
		/// <summary>
		/// Current assembly reference
		/// </summary>
		private readonly Assembly resourcesAssembly;
		
		/// <summary>
		/// List of compiled XSLTs
		/// (Note : it seems to contains only the oox2daisy.xsl compiled transform for now)
		/// </summary>
		private readonly Hashtable compiledProcessors;

		private readonly ArrayList MathEntities8879, MathEntities9573, MathMLEntities;

		private ArrayList fidilityLoss = new ArrayList();
		
		String LastErrorText = "";
		String LastValidatedFile = "";

		private readonly ChainResourceManager resourceManager;

		#region Fields
		public bool DirectTransform {
			set { this.isDirectTransform = value; }
			get { return this.isDirectTransform; }
		}

		public ArrayList SkipedPostProcessors {
			set { this.skipedPostProcessors = value; }
		}

		public string ExternalResources {
			set { this.externalResource = value; }
			get { return this.externalResource; }
		}

		public ArrayList FidilityLoss {
			get {
				return fidilityLoss;
			}
		}

		private IDictionary<string,List<string> > lostElements;
		public IDictionary<string, List<string> > LostElements {
			get {
				return lostElements;
			}
		}

		private ArrayList documentLanguages;
		public ArrayList DocumentLanguages {
			get {
				return documentLanguages;
			}
		}


		/// <summary>
		/// Pull the chain of post processors for the direct conversion
		/// </summary>
		protected virtual string[] DirectPostProcessorsChain {
			get { return null; }
		}

		/// <summary>
		/// Pull the chain of post processors for the reverse conversion
		/// </summary>
		protected virtual string[] ReversePostProcessorsChain {
			get { return null; }
		}

		/// <summary>
		/// Pull an XmlUrlResolver for embedded resources
		/// </summary>
		private XmlUrlResolver ResourceResolver {
			get {
				if (this.ExternalResources == null) {
					return new EmbeddedResourceResolver(this.resourcesAssembly,
					  this.GetType().Namespace, this.DirectTransform);
				} else {
					return new SharedXmlUrlResolver(this.DirectTransform);
				}
			}
		}

		/// <summary>
		/// Pull the input xml document to the xsl transformation
		/// </summary>
		private XmlReader Source {
			get {
				XmlReaderSettings xrs = new XmlReaderSettings {
					// do not look for DTD
					DtdProcessing = DtdProcessing.Prohibit
				};
				if (this.ExternalResources == null) {
					xrs.XmlResolver = this.ResourceResolver;
					return XmlReader.Create(SOURCE_XML, xrs);
				} else {
					return XmlReader.Create(this.ExternalResources + "/" + SOURCE_XML, xrs);
				}
			}
		}


		/// <summary>
		/// Pull the xslt settings
		/// </summary>
		private XsltSettings XsltProcSettings {
			get {
				// Enable xslt 'document()' function
				return new XsltSettings(true, false);
			}
		}

		
		#endregion

		public WordToDTBookXMLTransform(
			System.Resources.ResourceManager customResourceManager = null
		) {
			this.resourcesAssembly = Assembly.GetExecutingAssembly();
			this.skipedPostProcessors = new ArrayList();
			this.compiledProcessors = new Hashtable();

			this._validationErrors = new List<ValidationError>();

			this.resourceManager = new ChainResourceManager();
			
			// Add a default resource managers (for common labels)
			this.resourceManager.Add(
				new System.Resources.ResourceManager("Daisy.SaveAsDAISY.DaisyConverterLib.resources.Labels",
				Assembly.GetExecutingAssembly()));

			// additionnal resource manager
			if(customResourceManager != null) {
				this.resourceManager.Add(customResourceManager);
			}

			// Load math and mathml entities
			MathEntities8879 = new ArrayList();
			MathEntities9573 = new ArrayList();
			MathMLEntities = new ArrayList();

			MathEntities8879.Add("isobox.ent");
			MathEntities8879.Add("isocyr1.ent");
			MathEntities8879.Add("isocyr2.ent");
			MathEntities8879.Add("isodia.ent");
			MathEntities8879.Add("isolat1.ent");
			MathEntities8879.Add("isolat2.ent");
			MathEntities8879.Add("isonum.ent");

			MathEntities8879.Add("isopub.ent");

			MathMLEntities.Add("mmlalias.ent");
			MathMLEntities.Add("mmlextra.ent");

			MathEntities9573.Add("isoamsa.ent");
			MathEntities9573.Add("isoamsb.ent");
			MathEntities9573.Add("isoamsc.ent");
			MathEntities9573.Add("isoamsn.ent");
			MathEntities9573.Add("isoamso.ent");
			MathEntities9573.Add("isoamsr.ent");
			MathEntities9573.Add("isogrk3.ent");
			MathEntities9573.Add("isomfrk.ent");
			MathEntities9573.Add("isomopf.ent");
			MathEntities9573.Add("isomscr.ent");
			MathEntities9573.Add("isotech.ent");

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
		/// Transform a single word document into DTBook XML using xslts
		/// (using new settings centralization classes)
		/// </summary>
		/// <param name="document">Info of the document to convert (including output of the current document conversion) </param>
		/// <param name="conversion">Settings of the conversion</param>
		public XmlDocument ConvertDocument(DocumentParameters document, ConversionParameters conversion, XmlDocument mergeTarget = null) {
			
			fidilityLoss = new ArrayList();

			// temporary files in memory
			string tempInputPath = Path.GetTempFileName();
			string tempOutputPath =  Path.GetTempFileName();

			string conversionOutputDirectory = Directory.GetParent(document.OutputPath).FullName;
			progressMessageIntercepted(this, new DaisyEventArgs("Transform " + document.CopyPath + " to DTBook XML "+ document.OutputPath));

			try {
				File.Copy(document.CopyPath, tempInputPath, true);
				File.SetAttributes(tempInputPath, FileAttributes.Normal);

				XmlReader source = null;
				XmlWriter writer = null;
				ZipResolver zipResolver = null;

				try {
					XslCompiledTransform xslt = this.Load(false);
					zipResolver = new ZipResolver(tempInputPath);

					XsltArgumentList parameters = new XsltArgumentList();
					parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTMessageEvent);
					parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(onXSLTProgressMessageEvent);

					DaisyClass val = new DaisyClass(
								document.InputPath,
								tempInputPath,
								conversionOutputDirectory,
								document.ListMathMl,
								zipResolver.Archive,
								conversion.PipelineOutput);
					parameters.AddExtensionObject("urn:Daisy", val);


					parameters.AddParam("outputFile", "", tempOutputPath);

					foreach (DictionaryEntry myEntry in conversion.ConversionParametersHash) {
						parameters.AddParam(myEntry.Key.ToString(), "", myEntry.Value.ToString());
					}
					// write result to conversion.TempOutputFile or a random temp output file
					
					XmlWriter finalWriter;
#if DEBUG
					StreamWriter streamWriter = new StreamWriter(
						tempOutputPath,
						true, System.Text.Encoding.UTF8
					) { AutoFlush = true };
					finalWriter = new XmlTextWriter(streamWriter);
					
					Debug.WriteLine("OUTPUT FILE : '" + tempOutputPath + "'");
#else
					finalWriter = new XmlTextWriter(tempOutputPath, System.Text.Encoding.UTF8);
#endif
					writer = GetWriter(finalWriter);
					

					source = this.Source;
					// Apply the transformation

					xslt.Transform(source, parameters, writer, zipResolver);
				} finally {
					if (writer != null)
						writer.Close();
					if (source != null)
						source.Close();
					if (zipResolver != null)
						zipResolver.Dispose();
				}


				if (document.OutputPath != null) {
					if (File.Exists(document.OutputPath)) {
						File.Delete(document.OutputPath);
					}
					File.Move(tempOutputPath, document.OutputPath);
					CopyCSSToDestinationfolder(document.OutputPath);
					// TODO: The handling of DTD and math needs to be cleanedup
					// For now, the transform set a system dtd and
					Int16 value = (Int16)document.OutputPath.LastIndexOf("\\");
					String tempStr = document.OutputPath.Substring(0, value);
					
					CopyDTDToDestinationfolder(document.OutputPath);
					CopyMATHToDestinationfolder(document.OutputPath);
					if (conversion.Validate) {
						validateXML(document.OutputPath);
					}
					// We need to change this method : it does not only delete the dtds files,
					// it updates the file dtds
					DeleteDTD(tempStr + "\\" + "dtbook-2005-3.dtd", document.OutputPath, conversion.ScriptPath != null);
					DeleteMath(tempStr, conversion.ScriptPath != null);
				}
			} finally {
				if (File.Exists(tempInputPath)) {
					try {
						File.Delete(tempInputPath);
					} catch (IOException) {
						Debug.Write("could not delete temporary input file " + tempInputPath);
					}
				}
			}
			return this.MergeDocumentInTarget(document, mergeTarget);
		}

		/// <summary>
		/// This function does the following action on the "filename" file : 
		/// - replace the system local dtbook-2005-3.dtd by the public one
		/// - add the correct namespace to the dtbook declaration
		///
		/// Notes : 
		/// - if the value parameter is set to false, the function is only doing a read and rewrite the file on itself
		/// - If no closing mml:math tag is found, the code removes 
		///   - chars 203 to 1120, 
		///   - an empty line before the dtbook tag 
		///   - every occurences of the mml namespace declaration
		/// </summary>
		/// <param name="fileDTD">dtd file to be removed from disk</param>
		/// <param name="fileName">path of the file to be updated</param>
		/// <param name="value">boolean flag - if true, tha actions are applied on the updated file</param>
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
		

		public void DeleteMath(String folder, bool value)
		{
			if (value)
			{
				DeleteFile(folder + "\\mathml2.DTD");
				DeleteFile(folder + "\\mathml2-qname-1.mod");
				Directory.Delete(folder + "\\iso8879", true);
				Directory.Delete(folder + "\\iso9573-13", true);
				Directory.Delete(folder + "\\mathml", true);
			}
		}

		public void DeleteFile(String file)
		{
			if (File.Exists(file))
			{
				File.Delete(file);
			}
		}


		private XmlWriter GetWriter(XmlWriter writer)
		{
			string[] postProcessors = this.DirectPostProcessorsChain;
			/*if (!this.isDirectTransform)
			{
				postProcessors = this.ReversePostProcessorsChain;
			}*/
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

		#region XSLTs and conversion events handling and redispatching
		/// <summary>
		/// Progress and Feedback listener functions
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e">contains the feedback message or null if the message is a progress event</param>
		public delegate void XSLTMessagesListener(object sender, EventArgs e);

		/// <summary>
		/// Progress messages redispatcher
		/// </summary>
		private event XSLTMessagesListener progressMessageIntercepted;
		/// <summary>
		/// Progress messages redispatcher for master document
		/// </summary>
		private event XSLTMessagesListener progressMessageInterceptedMaster;
		/// <summary>
		/// Feedback message redispatcher
		/// </summary>
		private event XSLTMessagesListener feedbackMessageIntercepted;
		/// <summary>
		/// Validation
		/// </summary>
		private event XSLTMessagesListener feedbackValidationIntercepted;

		/// <summary>
		/// Events handler that receives XSLT messages and redispatche them using the progress and feedback intercepted event launcher.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void onXSLTMessageEvent(object sender, XsltMessageEncounteredEventArgs e) {
			if (e.Message.StartsWith("progress:")) {
				progressMessageIntercepted?.Invoke(this, new DaisyEventArgs(e.Message));
			} else if (e.Message.StartsWith("translation.oox2Daisy.")) {
				fidilityLoss.Add(e.Message);
				feedbackMessageIntercepted?.Invoke(this, new DaisyEventArgs(e.Message));
			}
		}

		/// <summary>
		/// Secondary handler that only redispatch "progress" messages 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void onXSLTProgressMessageEvent(object sender, XsltMessageEncounteredEventArgs e) {
			if (e.Message.StartsWith("progress:")) {
				progressMessageInterceptedMaster?.Invoke(this, new DaisyEventArgs(e.Message));
			}
		}

		public void AddProgressMessageListener(XSLTMessagesListener listener)
		{
			progressMessageIntercepted += listener;
		}

		public void AddProgressMessageListenerMaster(XSLTMessagesListener listener)
		{
			progressMessageInterceptedMaster += listener;
		}

		public void AddFeedbackMessageListener(XSLTMessagesListener listener)
		{
			feedbackMessageIntercepted += listener;
		}

		public void AddFeedbackValidationListener(XSLTMessagesListener listener)
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

		#endregion

		#region Files copy to output folder
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
			CopyingAssemblyFile(
				Path.GetDirectoryName(outputFile) + "\\mathml2-qname-1.mod",
				"mathml2-qname-1.mod");
			CopyingAssemblyFile(
				Path.GetDirectoryName(outputFile) + "\\mathml2.DTD",
				"mathml2.DTD");

			for (int i = 0; i < MathEntities8879.Count; i++)
			{
				Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso8879");
				CopyingAssemblyFile(
					Path.GetDirectoryName(outputFile) + "\\iso8879\\" + MathEntities8879[i].ToString(),
					MathEntities8879[i].ToString());
			}

			for (int i = 0; i < MathEntities9573.Count; i++)
			{
				Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\iso9573-13");
				CopyingAssemblyFile(
					Path.GetDirectoryName(outputFile) + "\\iso9573-13\\" + MathEntities9573[i].ToString(),
					MathEntities9573[i].ToString());
			}

			for (int i = 0; i < MathMLEntities.Count; i++)
			{
				Directory.CreateDirectory(Path.GetDirectoryName(outputFile) + "\\mathml");
				CopyingAssemblyFile(
					Path.GetDirectoryName(outputFile) + "\\mathml\\" + MathMLEntities[i].ToString(),
					MathMLEntities[i].ToString());

			}
		}

		public void CopyingAssemblyFile(String destinationFile, String resourceFileName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			Stream stream = null;

			foreach (string name in asm.GetManifestResourceNames())
			{
				if (name.EndsWith(resourceFileName))
				{
					stream = asm.GetManifestResourceStream(name);
					break;
				}

			}

			StreamReader reader = new StreamReader(stream);
			StreamWriter writer = new StreamWriter(destinationFile);
			string data = reader.ReadToEnd();
			writer.Write(data);
			reader.Close();
			writer.Close();
		}
		#endregion

		#region XML validation

		
		/// <summary>
		/// Schematron text report
		/// </summary>
		private string schematronReport;
		private bool isValid;

		/// <summary>
		/// List of ValidationError for the last validation request
		/// </summary>
		private readonly List<ValidationError> _validationErrors = new List<ValidationError>();
		/// <summary>
		/// List of ValidationError (Read only public field)
		/// </summary>
		public List<ValidationError> ValidationErrors { get => _validationErrors; }
		public string SchematronReport { get => schematronReport;}
		public bool IsValid { get => isValid; }
		


		/// <summary>
		/// Validate an XML file against DTDs.
		/// 
		/// Results/errors are send to all feedbackValidationIntercepted listeners
		/// (exposed with the AddFeedbackValidationListener method).
		/// 
		/// </summary>
		/// <param name="fileToValidate">File to be validated</param>
		public void validateXML(String fileToValidate)
		{
			LastValidatedFile = fileToValidate;
			isValid = true;
			schematronReport = "";

			XmlTextReader xml = new XmlTextReader(fileToValidate);

			XmlReaderSettings settings = new XmlReaderSettings {
				ValidationType = ValidationType.DTD,
				DtdProcessing = DtdProcessing.Parse
			};
			settings.ValidationEventHandler += new ValidationEventHandler(onValidationEvent);
			XmlReader xsd = XmlReader.Create(xml,settings);
			
			try
			{

				ArrayList errTxt = new ArrayList();
				for (int i = 0; i <= 4; i++)
					errTxt.Add("");
				while (xsd.Read())
				{
					errTxt[4] = errTxt[3];
					errTxt[3] = errTxt[2];
					errTxt[2] = errTxt[1];
					errTxt[1] = errTxt[0];
					errTxt[0] = xsd.ReadString();
					LastErrorText = "";
					for (int i = 4; i >= 0; i--)
						LastErrorText = LastErrorText + errTxt[i].ToString() + " ";
					if (LastErrorText.Contains("\n"))
						LastErrorText = LastErrorText.Replace("\n", "");
					if (LastErrorText.Contains("\r"))
						LastErrorText = LastErrorText.Replace("\r", "");
					if (LastErrorText.Length > 100)
						LastErrorText = LastErrorText.Substring(0, 100);
				}
				xsd.Close();

				// Schematron report generator
				Stream schematronXslStream = null;
				Assembly asm = Assembly.GetExecutingAssembly();
				foreach (string name in asm.GetManifestResourceNames())
				{
					if (name.EndsWith("Schematron.xsl"))
					{
						schematronXslStream = asm.GetManifestResourceStream(name);
						break;
					}
				}

				XmlReader schematronXsl = XmlReader.Create(schematronXslStream);
				XPathDocument validatedDocument = new XPathDocument(fileToValidate);

				XslCompiledTransform trans = new XslCompiledTransform(true);
				trans.Load(schematronXsl);

				XmlTextWriter reportWriter = new XmlTextWriter(Path.GetDirectoryName(fileToValidate) + "\\report.txt", null);
				trans.Transform(validatedDocument, null, reportWriter);

				reportWriter.Close();
				schematronXsl.Close();

				StreamReader reader = new StreamReader(Path.GetDirectoryName(fileToValidate) + "\\report.txt");
				if (!reader.EndOfStream)
				{
					schematronReport = reader.ReadToEnd();
					//errors += reader.ReadToEnd();
					feedbackValidationIntercepted?.Invoke(this, new DaisyEventArgs(schematronReport));
				}
				reader.Close();

				if (File.Exists(Path.GetDirectoryName(fileToValidate) + "\\report.txt"))
				{
					File.Delete(Path.GetDirectoryName(fileToValidate) + "\\report.txt");
				}

				// Check whether the document is valid or invalid.
				if (isValid == false)
				{
					if (feedbackValidationIntercepted != null)
					{
						foreach (ValidationError found in ValidationErrors) {
							feedbackValidationIntercepted(
								this, new DaisyEventArgs(
									" Line Number:Position : " + found.error.Exception.LineNumber + ":" + found.error.Exception.LinePosition + Environment.NewLine +
									" Message : " + found.error.Message + Environment.NewLine + 
									" Reference Text :  " + found.referenceText + Environment.NewLine));
						}
						
					}
				}
			}
			catch (UnauthorizedAccessException a)
			{
				xsd.Close();
				//dont have access permission
				feedbackValidationIntercepted?.Invoke(this, new DaisyEventArgs(a.Message));
			}
			catch (Exception a)
			{
				xsd.Close();
				//and other things that could go wrong
				feedbackValidationIntercepted?.Invoke(this, new DaisyEventArgs(a.Message));
			}
		}


		/// <summary>
		/// XML Validation events callback
		/// Add the validation event to the error stack
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="args"></param>
		public void onValidationEvent(object sender, ValidationEventArgs args)
		{
			isValid = false;
			_validationErrors.Add(new ValidationError(args, LastValidatedFile, LastErrorText));
			/*errors += " Line Number : " + args.Exception.LineNumber + " and " +
			 " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
			 " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;*/
		}

		#endregion XML validation

		#region Merging document

		// Counter of merged document to be written as footnote
		private int mergedDocumentCounter = 0;
		// MathML data extracted
		private string cutData = "";

		private string mergeErrors = "";

		public string MergeErrors { get => mergeErrors; }

		/// <summary>
		/// Merge a document with a previous document. <br/>
		/// The document content is merged either on the position of the subdocument in the master document (using the ResourceId)
		/// or at the end of the merge target body matter.
		/// </summary>
		/// <param name="documentToMerge">Document to merge in the target</param>
		/// <param name="mergeTarget">XmlDocument of previous merge operations, should be null if no previous merge operation has been done</param>
		/// <returns>The merged xml document, or the xml result of the conversion if no or an empty target is used as input)</returns>
		public XmlDocument MergeDocumentInTarget(
			DocumentParameters documentToMerge,
			XmlDocument mergeTarget
		) {

			XmlDocument tempDoc = new XmlDocument();
			cutData = mergeTarget == null ? "" : cutData;
			if (mergeTarget != null) { // An existing merge target (empty or not) is requested
				ReplaceData(documentToMerge.OutputPath, out cutData);
			}
			
			tempDoc.Load(documentToMerge.OutputPath);
			if (mergeTarget == null || mergeTarget.ChildNodes.Count == 0) {
				// No (or empty) merge target,
				// Prepare document merging data 
				// (document languages and parts of documents that are not kept, like TOCs)
				// Also keep a counter of document merged to add a footnote
				documentLanguages = new ArrayList();
				lostElements = new Dictionary<string, List<string> >();
				mergedDocumentCounter = 0;
				mergeErrors = "";

				// return the current xml document (that will remain on disk)
				return tempDoc;
			} else try { // merge the document in the target and delete it
					progressMessageIntercepted?.Invoke(this, new DaisyEventArgs("Merging document " + (mergedDocumentCounter + 1) + " with previous result - " + documentToMerge.InputPath));
					XmlNode tempNode = null;

					tempDoc = SetFootnote(tempDoc, "subDoc" + (mergedDocumentCounter + 1));

					for (int i = 0; i < tempDoc.SelectSingleNode("//head").ChildNodes.Count; i++) {
						tempNode = tempDoc.SelectSingleNode("//head").ChildNodes[i];

						if (tempNode.Attributes[0].Value == "dc:Language") {
							if (!documentLanguages.Contains(tempNode.Attributes[1].Value)) {
								documentLanguages.Add(tempNode.Attributes[1].Value);
							}
						}
					}

					for (int i = 0; i < tempDoc.SelectSingleNode("//bodymatter").ChildNodes.Count; i++) {
						tempNode = tempDoc.SelectSingleNode("//bodymatter").ChildNodes[i];

						if (tempNode != null) {
							XmlNode addBodyNode = mergeTarget.ImportNode(tempNode, true);
							if (addBodyNode != null) {
								if(documentToMerge.ResourceId != null) { // Subdocument being reinserted
									XmlNode subDocNode = mergeTarget.SelectSingleNode(
										"//subdoc[@rId='" + documentToMerge.ResourceId + "']"
									);
									subDocNode.ParentNode.InsertBefore(
										addBodyNode,
										subDocNode
									);
								} else { // Document being appended to body matter ?
									mergeTarget.LastChild.LastChild.FirstChild.NextSibling.AppendChild(addBodyNode);
								}
							}
							   
						}
					}

					tempNode = tempDoc.SelectSingleNode("//frontmatter/level1[@class='print_toc']");
					if (tempNode != null) {
						if(lostElements[documentToMerge.InputPath] == null) {
							lostElements[documentToMerge.InputPath] = new List<string>();
						}
						if(!lostElements[documentToMerge.InputPath].Contains("TOC not translated")) {
							lostElements[documentToMerge.InputPath].Add("TOC not translated");
						}
					}
					if(documentToMerge.ResourceId != null ) {
						mergeTarget.SelectSingleNode(
							"//subdoc[@rId='" + documentToMerge.ResourceId + "']"
						).ParentNode.RemoveChild(
							mergeTarget.SelectSingleNode(
								"//subdoc[@rId='" + documentToMerge.ResourceId + "']"
							)
						);

					}
					

					XmlNode node = tempDoc.SelectSingleNode("//rearmatter");

					if (node != null) {
						for (int i = 0; i < tempDoc.SelectSingleNode("//rearmatter").ChildNodes.Count; i++) {
							tempNode = tempDoc.SelectSingleNode("//rearmatter").ChildNodes[i];

							if (tempNode != null) {
								XmlNode addRearNode = mergeTarget.ImportNode(tempNode, true);
								if (addRearNode != null)
									mergeTarget.LastChild.LastChild.LastChild.AppendChild(addRearNode);
							}
						}
					}
					mergedDocumentCounter++;
				} catch (Exception e) {
					mergeErrors = mergeErrors + "\n" + " \"" + documentToMerge.InputPath + "\"";
					mergeErrors = mergeErrors + "\n" + "Validation error:" + "\n" + e.Message + "\n";
				}
			// delete document that have been merged in the target
			if (File.Exists(documentToMerge.OutputPath)) {
				File.Delete(documentToMerge.OutputPath);
			}

			return mergeTarget;
		}


		// 
		/// <summary>
		/// Finalize and save a merged document
		/// </summary>
		/// <param name="mergeTarget"></param>
		/// <param name="conversion"></param>
		public void finalizeAndSaveMergedDocument(
			XmlDocument mergeTarget,
			ConversionParameters conversion
		) {
			progressMessageIntercepted?.Invoke(this, new DaisyEventArgs("Cleaning and saving document to " + conversion.OutputPath));
			string outputDirectory = conversion.OutputPath.EndsWith(".xml") ?
					Directory.GetParent(conversion.OutputPath).FullName :
					conversion.OutputPath;

			SetPageNum(mergeTarget);
			SetImage(mergeTarget);
			SetLanguage(mergeTarget, documentLanguages);
			RemoveSubDoc(mergeTarget);
			mergeTarget.Save(conversion.OutputPath);
			ReplaceData(conversion.OutputPath, true, cutData);
			CopyDTDToDestinationfolder(conversion.OutputPath);
			CopyMATHToDestinationfolder(conversion.OutputPath);
			if(conversion.Validate) validateXML(conversion.OutputPath);
			ReplaceData(conversion.OutputPath, false, cutData);
			if (File.Exists(outputDirectory + "\\dtbook-2005-3.dtd")) {
				File.Delete(outputDirectory + "\\dtbook-2005-3.dtd");
			}
			DeleteMath(outputDirectory, true);
		}


		public XmlDocument SetFootnote(XmlDocument mergeXmlDoc, String SubDocFootnum) {
			int footnoteCount = 1, endnoteCount = 1;
			XmlNodeList noteList = mergeXmlDoc.SelectNodes("//note");
			if (noteList != null) {
				for (int i = 1; i <= noteList.Count; i++) {
					if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Footnote") {
						mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "footnote-" + footnoteCount.ToString();
						footnoteCount++;
					}
					if (mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[1].InnerText == "Endnote") {
						mergeXmlDoc.SelectNodes("//note").Item(i - 1).Attributes[0].InnerText = SubDocFootnum + "endnote-" + endnoteCount.ToString();
						endnoteCount++;
					}
				}
			}

			footnoteCount = 1;
			endnoteCount = 1;
			noteList = mergeXmlDoc.SelectNodes("//noteref");
			if (noteList != null) {

				for (int i = 1; i <= noteList.Count; i++) {
					if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Footnote") {
						mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "footnote-" + footnoteCount.ToString();
						footnoteCount++;
					}
					if (mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[1].InnerText == "Endnote") {
						mergeXmlDoc.SelectNodes("//noteref").Item(i - 1).Attributes[0].InnerText = "#" + SubDocFootnum + "endnote-" + endnoteCount.ToString();
						endnoteCount++;
					}
				}
			}

			return mergeXmlDoc;

		}

		/* Function which creates unique ID to page numbers*/
		public void SetPageNum(XmlDocument mergeXmlDoc) {
			XmlNodeList pageList = mergeXmlDoc.SelectNodes("//pagenum");
			for (int i = 1; i <= pageList.Count; i++) {
				mergeXmlDoc.SelectNodes("//pagenum").Item(i - 1).Attributes[1].InnerText = "page" + i.ToString();
			}
		}

		/* Function which creates unique ID to Images*/
		public void SetImage(XmlDocument mergeXmlDoc) {
			XmlNodeList imageList = mergeXmlDoc.SelectNodes("//img");
			int j = 0;
			for (int i = 1; i <= imageList.Count; i++) {
				if (mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText.StartsWith("rId")) {
					mergeXmlDoc.SelectNodes("//img").Item(i - 1).Attributes[0].InnerText = "rId" + j.ToString();
					j++;
				}
			}
			XmlNodeList captionList = mergeXmlDoc.SelectNodes("//caption");
			for (int i = 1; i <= captionList.Count; i++) {
				XmlNode prevNode = mergeXmlDoc.SelectNodes("//caption").Item(i - 1).PreviousSibling;
				if (prevNode != null) {
					String rId = prevNode.Attributes[0].InnerText;
					mergeXmlDoc.SelectNodes("//caption").Item(i - 1).Attributes[0].InnerText = rId;
				}
			}
		}

		/* Function which creates language info of all sub documents in master.xml*/
		public void SetLanguage(XmlDocument mergeXmlDoc, ArrayList mergingLanguagesList) {
			XmlNodeList languageList = mergeXmlDoc.SelectNodes("//meta[@name='dc:Language']");

			for (int i = 0; i < languageList.Count; i++) {
				if (mergingLanguagesList.Contains(languageList[i].Attributes[1].Value)) {
					int indx = mergingLanguagesList.IndexOf(languageList[i].Attributes[1].Value);
					mergingLanguagesList.RemoveAt(indx);
				}
			}

			for (int i = 0; i < mergingLanguagesList.Count; i++) {
				XmlElement tempLang = mergeXmlDoc.CreateElement("meta");
				tempLang.SetAttribute("name", "dc:Language");
				tempLang.SetAttribute("content", mergingLanguagesList[i].ToString());
				mergeXmlDoc.SelectNodes("//head").Item(0).AppendChild(tempLang);
			}
		}

		/* Function which removes subdoc elements from the master.xml*/
		public void RemoveSubDoc(XmlDocument mergeXmlDoc) {
			XmlNodeList subDocList = mergeXmlDoc.SelectNodes("//subdoc");
			if (subDocList != null) {
				for (int i = 0; i < subDocList.Count; i++) {
					subDocList.Item(i).ParentNode.RemoveChild(subDocList.Item(i));
				}
			}
		}

		/// <summary>
		/// This method remove the doctype
		/// </summary>
		/// <param name="fileName"></param>
		/// <param name="cutData"></param>
		public void ReplaceData(String fileName, out string cutData) {
			StreamReader reader = new StreamReader(fileName);
			string data = reader.ReadToEnd();
			reader.Close();
			
			cutData = "";
			StreamWriter writer = new StreamWriter(fileName);
			bool hasMathMl = data.Contains("</mml:math>");
			string doctypeToRemove = !hasMathMl ?
				"<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd' >" :
				"<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'[<!ENTITY % MATHML.prefixed \"INCLUDE\" ><!ENTITY % MATHML.prefix \"mml\"><!ENTITY % Schema.prefix \"sch\"><!ENTITY % XLINK.prefix \"xlp\"><!ENTITY % MATHML.Common.attrib \"xlink:href    CDATA       #IMPLIED xlink:type     CDATA       #IMPLIED   class          CDATA       #IMPLIED  style          CDATA       #IMPLIED  id             ID          #IMPLIED  xref           IDREF       #IMPLIED  other          CDATA       #IMPLIED   xmlns:dtbook   CDATA       #FIXED 'http://www.daisy.org/z3986/2005/dtbook/' dtbook:smilref CDATA       #IMPLIED\"><!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>%mathML2;<!ENTITY % externalFlow \"| mml:math\"><!ENTITY % externalNamespaces \"xmlns:mml CDATA #FIXED 'http://www.w3.org/1998/Math/MathML'\">]>";
			// Remove doctype
			data = data.Replace(
				doctypeToRemove,
				"<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>"
			);
			if (hasMathMl) { // I don't know what this part extracts exactly
				cutData = data.Substring(95, 1091);
				data = data.Remove(95, 1091);
			}

			// Remove namespace and add mathml if needed
			// (i assumed, but the previous codebase add the namespace
			// if no mml:math closing tag is found
			data = data.Replace(
					"<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"",
					"<dtbook version=\"" + "2005-3\"" + (!hasMathMl ? " xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"" : ""));

			writer.Write(data);
			writer.Close();
		}

		/* Function which merges subdocument.xml and master.xml*/
		public void ReplaceData(String fileName, bool value, in string cutData = "") {
			StreamReader reader = new StreamReader(fileName);
			string data = reader.ReadToEnd();
			reader.Close();
			string tempData = "";
			StreamWriter writer = new StreamWriter(fileName);
			if (value) {
				if (!data.Contains("</mml:math>")) {
					data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?><!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>");
					data = data.Replace("<dtbook version=\"" + "2005-3\" xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xml:lang=", "<dtbook version=\"" + "2005-3\" xml:lang=");
				} else {
					tempData = cutData.Replace("<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'", "<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'");
					tempData = tempData.Replace("<!ENTITY % mathML2 PUBLIC \"-//W3C//DTD MathML 2.0//EN\" \"http://www.w3.org/Math/DTD/mathml2/mathml2.dtd\">", "<!ENTITY % mathML2 SYSTEM 'mathml2.dtd'>");
					data = data.Replace("<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>", "<?xml-stylesheet href=\"dtbookbasic.css\" type=\"text/css\"?>" + tempData);
				}
			} else {
				if (!data.Contains("</mml:math>")) {
					data = data.Replace("<!DOCTYPE dtbook SYSTEM 'dtbook-2005-3.dtd'>", "<!DOCTYPE dtbook PUBLIC '-//NISO//DTD dtbook 2005-3//EN' 'http://www.daisy.org/z3986/2005/dtbook-2005-3.dtd'>");
					data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
				} else {
					data = data.Replace(tempData, cutData);
					data = data.Replace("<dtbook version=\"" + "2005-3\"", "<dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\"");
				}
			}
			writer.Write(data);
			writer.Close();
		}
		#endregion

	}


}
