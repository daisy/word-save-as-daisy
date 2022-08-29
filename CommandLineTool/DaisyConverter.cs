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
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Schema;
using System.Xml.XPath;
using System.Xml.Xsl;
using Daisy.SaveAsDAISY.Conversion;
using System.Xml;
using System.Resources;
using System.IO.Packaging;
using System.Drawing.Imaging;

using MSword = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Addins.Word2007;
using Daisy.SaveAsDAISY.Forms;

namespace Daisy.SaveAsDAISY.CommandLineTool {
    enum ControlType : int {
        CTRL_C_EVENT = 0,
        CTRL_BREAK_EVENT = 1,
        CTRL_CLOSE_EVENT = 2,
        CTRL_LOGOFF_EVENT = 5,
        CTRL_SHUTDOWN_EVENT = 6
    };

    enum Direction {
        None,
        DocxToXml,
        BatchDocx,
        BatchOwnDocx
    };

    enum Label {
        ERROR,
        WARNING,
        INFO,
        DEBUG
    };

    delegate int ControlHandlerFonction(ControlType control);

    

    

    /// <summary>
    /// DaisyTranslatorTest is a CommandLine Program to test the conversion
    /// of an OpenXML file into a Daisy format file.
    /// 
    /// Execute the command without argument to see the options.
    /// </summary>
    public class DaisyConverter {
        private string input = null;                     // input path
        private string titleProp = null;
        private string output = null;                    // output path
        private bool open = false;                       // try to open the result of the transformations
        private bool preprocessOnly = false;
        private string reportPath = null;                // file to save report
        private int reportLevel = Report.INFO_LEVEL;     // file to save report
        private string xslPath = null;                   // Path to an external stylesheet
        private ArrayList skipedPostProcessors = null;   // Post processors to skip (identified by their names)
        private Direction transformDirection = Direction.DocxToXml; // direction of conversion
        public Report report = null;
        private Word word = null;
        private OoxValidator ooxValidator = null;
        private Direction batch = Direction.None;
        Hashtable myHT = new Hashtable();
        Hashtable myLabel = new Hashtable();
        public ArrayList fidilityLoss = null;
        public ArrayList fidilityLossMsg;
        private static bool isValid;
        String[] listDoc;
        String wildCard = "None";
        ArrayList batchDoc = new ArrayList();
        XmlDocument mergeXmlDoc;
        ArrayList openSubdocs;
        PackageRelationship packRelationship = null;
        ArrayList lostElements = new ArrayList();
        ArrayList MathList8879, MathList9573, MathListmathml;
        String errorText = "";
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        
        private Pipeline1 _postprocessingPipeline = null;
        public Pipeline1 PostprocessingPipeline {
            get {
                if (ConverterHelper.PipelineIsInstalled() && _postprocessingPipeline == null)
                    _postprocessingPipeline = Pipeline1.Instance;
                return _postprocessingPipeline;
            }
            set => _postprocessingPipeline = value;
        }


#if MONO
		static bool SetConsoleCtrlHandler(ControlHandlerFonction handlerRoutine, bool add) 
		{ return true; }
#else
        [DllImport("kernel32")]
        static extern bool SetConsoleCtrlHandler(ControlHandlerFonction handlerRoutine, bool add);
#endif

        int MyHandler(ControlType control) {
            if (word != null) {
                try {
                    word.Quit();
                } catch {
                    Console.WriteLine("Unable to close Word, please close it manually");
                }
            }
            return 0;
        }

        /// <summary>
        /// Progam entry point
        /// </summary>
        /// <see cref="usage"/>
        /// <param name="args">Command Line arguments </param>
        [STAThread]
        public static void Main(String[] args) {
            DaisyConverter tester = new DaisyConverter();
            Hashtable myHT = new Hashtable();
            Hashtable myLabel = new Hashtable();

            ControlHandlerFonction myHandler = new ControlHandlerFonction(tester.MyHandler);
            SetConsoleCtrlHandler(myHandler, true);
            try {
                tester.ParseCommandLine(args);
            } catch (Exception e) {
                Environment.ExitCode = 1;
                Console.WriteLine("Error when parsing command line: " + e.Message);
                usage();
                return;
            }
            try {
                tester.Proceed();
                Console.WriteLine("Done.");
            } catch (Exception ex) {
                Console.WriteLine("Exception raised when running test : " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }


        /*Constructor*/
        private DaisyConverter() {
            this.skipedPostProcessors = new ArrayList();
            myLabel.Add("translation.oox2Daisy.commentReference", " Comment Reference is not translated");
            myLabel.Add("translation.oox2Daisy.ExternalImage", " External Images are not translated");
            myLabel.Add("translation.oox2Daisy.GeometricShapeElement", " Geometric shapes is is not translated");
            myLabel.Add("translation.oox2Daisy.ImageContent", " Image is Corrupted");
            myLabel.Add("translation.oox2Daisy.object", " Object is not translated");
            myLabel.Add("translation.oox2Daisy.sdtElement", " is not translated");
            myLabel.Add("translation.oox2Daisy.shapeElement", " Shapes are not translated");
            myLabel.Add("translation.oox2Daisy.UncoveredElement", " is not translated");
            myLabel.Add("translation.oox2Daisy.ImageCaption", "Caption above the Image is not translated");
            myLabel.Add("translation.oox2Daisy.TableCaption", "Caption below the Table is not translated");
            myLabel.Add("translation.oox2Daisy.Image", "Images are not translated");


        }

        /* Function which takes the arguements from user and it validates arguements*/
        private void ParseCommandLine(string[] args) {
            this.RetrieveArgs(args);
            this.CheckPaths();
        }

        /// <summary>
        /// Parses the command line arguments and assigns the values to the corresponding variables
        /// </summary>
        /// <param name="args"></param>
        private void RetrieveArgs(string[] args) {
            for (int i = 0; i < args.Length; i++) {
                switch (args[i]) {
                    case "/I":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Input missing");
                        } else if (args[i].Contains("*")) {
                            if (File.Exists(this.input)) {
                                throw new DaisyCommandLineException("Input file does not exists");
                            }
                            Int16 wc1 = Int16.Parse(args[i].LastIndexOf('\\').ToString());
                            String str1 = args[i].Substring(0, wc1);
                            Int16 wc2 = Int16.Parse(args[i].LastIndexOf('*').ToString());
                            String str = args[i].Substring(wc1 + 1, wc2 - wc1 - 1);
                            Int16 wc3 = Int16.Parse(args[i].LastIndexOf('.').ToString());
                            String str2 = args[i].Substring(wc2 + 1, wc3 - wc2 - 1);

                            if (str.StartsWith(str) && str2.EndsWith(str2)) {
                                listDoc = Directory.GetFiles(str1 + '\\', str + '*' + str2 + ".docx", SearchOption.TopDirectoryOnly);
                                this.input = Path.GetDirectoryName(args[i]);
                                this.batch = Direction.DocxToXml;
                                wildCard = "*";

                                if (listDoc.Length == 0) {
                                    throw new DaisyCommandLineException("Input file does not Exists");
                                }

                            }
                        } else if (args[i].Contains("?")) {
                            Int16 wc1 = Int16.Parse(args[i].LastIndexOf('\\').ToString());
                            String str1 = args[i].Substring(0, wc1);
                            Int16 wc2 = Int16.Parse(args[i].LastIndexOf('?').ToString());
                            String str2 = args[i].Substring(wc1 + 1, wc2 - wc1 - 1);
                            Int16 wc3 = Int16.Parse(args[i].LastIndexOf('.').ToString());
                            String str3 = args[i].Substring(wc2 + 1, wc3 - wc2 - 1);
                            if (str2.StartsWith(str2) && str3.EndsWith(str3)) {
                                listDoc = Directory.GetFiles(str1 + '\\', str2 + '?' + str3 + ".docx", SearchOption.TopDirectoryOnly);
                                this.input = Path.GetDirectoryName(args[i]);
                                this.batch = Direction.DocxToXml;
                                wildCard = "?";


                                if (listDoc.Length == 0) {
                                    throw new DaisyCommandLineException("Input file does not Exists");
                                }


                            }
                        } else {
                            this.input = args[i];
                            if (Directory.Exists(this.input)) {
                                String[] filesBatch = Directory.GetFiles(this.input);
                                foreach (string input in filesBatch) {
                                    if (!Path.GetFileNameWithoutExtension(input).StartsWith("~$") && input.EndsWith(".docx"))
                                        batchDoc.Add(input);
                                }

                            }
                        }
                        break;
                    case "/O":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Output missing");
                        }
                        if (!args[i].ToLower().EndsWith(".xml")) {

                            if (this.input.ToLower().EndsWith(".docx")) {
                                int length = this.input.LastIndexOf("\\");
                                string s1 = this.input.Substring(length);
                                string s2 = s1.ToLower().Replace(".docx", ".xml");
                                this.output = args[i] + s2;
                            } else
                                this.output = args[i];
                        } else if (args[i].ToLower().EndsWith(".xml")) {
                            this.output = args[i];
                        } else {
                            this.output = args[i];
                        }

                        break;
                    case "/CREATOR":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Creator missing");
                        }
                        myHT.Add("Creator", args[i]);
                        break;
                    case "/TITLE":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Title missing");
                        }
                        this.titleProp = args[i];
                        myHT.Add("Title", args[i]);
                        break;
                    case "/PUBLISHER":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Publisher missing");
                        }
                        myHT.Add("Publisher", args[i]);
                        break;
                    case "/UID":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Uid missing");
                        }
                        myHT.Add("UID", args[i]);
                        break;
                    case "/BATCH-DOCX":
                        this.batch = Direction.DocxToXml;
                        break;
                    case "/REPORT":
                        if (++i == args.Length) {
                            throw new DaisyCommandLineException("Report file missing");
                        }
                        this.reportPath = args[i];
                        break;
                    case "/M":
                        if (File.Exists(this.input)) {
                            this.batch = Direction.BatchOwnDocx;
                            myHT.Add("MasterSub", "Yes");
                        } else {
                            this.batch = Direction.BatchDocx;
                        }
                        break;
                    case "/CPAGE":
                        if (!myHT.Contains("Custom"))
                            myHT.Add("Custom", "Custom");
                        else
                            throw new DaisyCommandLineException("Choose one PageNumber Style");
                        break;
                    case "/APAGE":
                        if (!myHT.Contains("Custom"))
                            myHT.Add("Custom", "Automatic");
                        else
                            throw new DaisyCommandLineException("Choose one PageNumber Style");
                        break;
                    case "/STYLE":
                        myHT.Add("CharacterStyles", "True");
                        break;
                    case "/PREPROCESS":
                        preprocessOnly = true;
                        break;
                    default:
                        break;
                }
            }

            if (!myHT.Contains("Custom"))
                myHT.Add("Custom", "Custom");
            if (!myHT.Contains("CharacterStyles"))
                myHT.Add("CharacterStyles", "False");
            if (!myHT.Contains("ImageSizeOption"))
                myHT.Add("ImageSizeOption", "original");
        }

        /// <summary>
        /// Convert a docx file
        /// </summary>
        private void Proceed() {
            this.report = new Report(this.reportPath, this.reportLevel);
            String titleDoc = "";
            switch (this.batch) {
                case Direction.DocxToXml:
                    this.ProceedBatchDaisy();
                    break;
                case Direction.BatchDocx:
                    titleDoc = DocPropTitle(batchDoc[0].ToString());
                    if (this.titleProp == null && titleDoc == "") {
                        Environment.ExitCode = 1;
                        Console.WriteLine("Error when parsing command line: Title is Missing for " + Path.GetFileName(batchDoc[0].ToString()));
                        usage();
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        this.ProceedBatchDocx();
                    }
                    break;
                case Direction.BatchOwnDocx:
                    titleDoc = DocPropTitle(this.input);
                    if (this.titleProp == null && titleDoc == "") {
                        Environment.ExitCode = 1;
                        Console.WriteLine("Error when parsing command line: Title is Missing");
                        usage();
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        //this.ProceedOwnBatchDocx();
                        this.ProceedSingleFile(this.input, this.output, this.transformDirection, myHT);
                    }
                    break;
                default:  // no batch mode
                    // instanciate word if needed
                    if (this.transformDirection == Direction.DocxToXml && this.open) {
                        word = new Word();
                        word.Visible = false;
                    }
                    titleDoc = DocPropTitle(this.input);
                    if (this.titleProp == null && titleDoc == "") {
                        Environment.ExitCode = 1;
                        Console.WriteLine("Error when parsing command line: Title is Missing");
                        usage();
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        this.ProceedSingleFile(this.input, this.output, this.transformDirection, myHT);
                    }

                    // close word if needed
                    if (this.open) {
                        word.Quit();
                    }
                    break;
            }

            this.report.Close();
        }

        /* Core Function to translate the current document */
        private void ProceedSingleFile(string input, string output, Direction transformDirection, Hashtable table) {
            bool converted = false;
            bool validated = false;
            validated = ValidateFile(input);
            if (validated) {

                object _input = (object)input;
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                MSword.Document doc = app.Documents.Open(ref _input);

                GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(app);
                Conversion.Script pipelineScript = this.PostprocessingPipeline?.getScript("_postprocess");

                ConversionParameters conversion = new ConversionParameters(app.Version, pipelineScript);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);
                DocumentParameters currentDocument = converter.PreprocessDocument(app.ActiveDocument.FullName);
                if (converter.requestUserParameters(currentDocument) == ConversionStatus.ReadyForConversion) {
                    ConversionResult result = converter.Convert(currentDocument);
                    converted = result.Succeeded;
                } else {
                    eventsHandler.onConversionCanceled();
                }

            }

            if (!converted) {
                Environment.ExitCode = 1;

            }

        }


        /*Function which translates a batch of docx files*/
        private void ProceedBatchDaisy() {
            string ext;
            string targetExt;
            switch (this.batch) {
                case Direction.DocxToXml:
                    ext = "docx";
                    targetExt = ".xml";
                    break;
                default:
                    throw new ArgumentException("wrong batch mode");
            }
            string[] files;
            if (wildCard == "*") {

                files = listDoc;
                foreach (string input in files) {
                    string output = this.GenerateOutputName(this.output, input, targetExt);
                    String titleDoc = DocPropTitle(input);
                    if (this.titleProp == null && titleDoc == "") {
                        this.report.AddLog(input, "Document is not Translated. Title is missing", Report.INFO_LEVEL);
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        this.ProceedSingleFile(input, output, this.batch, myHT);
                    }
                }
            } else if (wildCard == "?") {

                files = listDoc;
                foreach (string input in files) {
                    string output = this.GenerateOutputName(this.output, input, targetExt);
                    String titleDoc = DocPropTitle(input);
                    if (this.titleProp == null && titleDoc == "") {
                        this.report.AddLog(input, "Document is not Translated. Title is missing", Report.INFO_LEVEL);
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        this.ProceedSingleFile(input, output, this.batch, myHT);
                    }
                }
            } else {
                SearchOption option = SearchOption.TopDirectoryOnly;
                files = Directory.GetFiles(this.input, "*." + ext, option);
                foreach (string input in files) {
                    string output = this.GenerateOutputName(this.output, input, targetExt);
                    String titleDoc = DocPropTitle(input);
                    if (this.titleProp == null && titleDoc == "") {
                        this.report.AddLog(input, "Document is not Translated. Title is missing", Report.INFO_LEVEL);
                    } else {
                        if (!myHT.ContainsKey("Title")) {
                            if (this.titleProp == null)
                                myHT.Add("Title", titleDoc);
                        }
                        this.ProceedSingleFile(input, output, this.batch, myHT);
                    }
                }
            }

        }

        /*Function which translates a batch of docx files*/
        private void ProceedBatchDocx() {
            try {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                GraphicalEventsHandler eventsHandler = new GraphicalEventsHandler();
                IDocumentPreprocessor preprocess = new DocumentPreprocessor(app);

                Conversion.Script pipelineScript = this.PostprocessingPipeline?.getScript("_postprocess");

                ConversionParameters conversion = new ConversionParameters(app.Version, pipelineScript);
                WordToDTBookXMLTransform documentConverter = new WordToDTBookXMLTransform();
                GraphicalConverter converter = new GraphicalConverter(preprocess, documentConverter, conversion, eventsHandler);


                List<string> documentsPathes = new List<string>();
                foreach (string docPath in batchDoc) {
                    documentsPathes.Add(docPath);
                }
                if (documentsPathes != null && documentsPathes.Count > 0) {
                    List<DocumentParameters> documents = new List<DocumentParameters>();

                    foreach (string inputPath in documentsPathes) {
                        DocumentParameters subDoc = null;
                        try {
                            subDoc = converter.PreprocessDocument(inputPath);
                        } catch (Exception e) {
                            string errors = "Convertion aborted due to the following errors found while preprocessing " + inputPath + ":\r\n" + e.Message;
                            eventsHandler.onPreprocessingError(inputPath, errors);
                        }
                        if (subDoc != null) {
                            documents.Add(subDoc);
                        } else {
                            // abort documents conversion
                            documents.Clear();
                            break;
                        }
                    }
                    if (documents.Count > 0) converter.Convert(documents);
                }
            } catch (Exception e) {
                this.report.AddLog("", "Erros in bacch conversion:\r\n" + e.Message, Report.ERROR_LEVEL);
            }
        }

        #region Validation

        /*Core Function to validate the input document*/
        private bool ValidateFile(string input) {
            try {
                if (this.ooxValidator == null) {
                    this.report.AddComment("Instanciating validator...");
                    this.ooxValidator = new OoxValidator(this.report);
                }
                this.ooxValidator.validate(input);
                this.report.AddComment("**********************");
                this.report.AddLog(input, "Input file (" + input + ") is valid", Report.INFO_LEVEL);
                return true;
            } catch (OoxValidatorException e) {
                this.report.AddComment("**********************");
                this.report.AddLog(input, "Input file (" + input + ") is invalid", Report.WARNING_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            } catch (Exception e) {
                this.report.AddLog(input, "An unexpected exception occured when trying to validate " + input, Report.ERROR_LEVEL);
                this.report.AddLog(input, e.Message + "(" + e.StackTrace + ")", Report.DEBUG_LEVEL);
                return false;
            }

        }

        /*Function to validate the input document*/
        public void ValidateOutputFile(String outFile) {
            isValid = true;
            XmlTextReader xml = new XmlTextReader(outFile);
            XmlValidatingReader xsd = new XmlValidatingReader(xml);

            try {
                xsd.ValidationType = ValidationType.DTD;
                xsd.ValidationEventHandler += new ValidationEventHandler(MyValidationEventHandler);

                while (xsd.Read()) {
                    errorText = xsd.ReadString();
                    if (errorText.Length > 100)
                        errorText = errorText.Substring(0, 100);
                }
                xsd.Close();

                Stream stream = null;
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames()) {
                    if (name.EndsWith("Shematron.xsl")) {
                        stream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }

                XmlReader rdr = XmlReader.Create(stream);
                XPathDocument doc = new XPathDocument(outFile);

                XslCompiledTransform trans = new XslCompiledTransform(false);
                trans.Load(rdr);

                XmlTextWriter myWriter = new XmlTextWriter(Path.GetDirectoryName(outFile) + "\\report.txt", null);
                trans.Transform(doc, null, myWriter);

                myWriter.Close();
                rdr.Close();

                StreamReader reader = new StreamReader(Path.GetDirectoryName(outFile) + "\\report.txt");
                if (!reader.EndOfStream) {
                    String error = reader.ReadToEnd();
                    report.AddLog(input, "Validation Error of converted DAISY File: \n" + error, Report.ERROR_LEVEL);
                    if (isValid == true) {
                        report.AddLog(input, "Translated DAISY file is not Valid: ", Report.ERROR_LEVEL);
                    }
                }
                reader.Close();

                if (File.Exists(Path.GetDirectoryName(outFile) + "\\report.txt")) {
                    File.Delete(Path.GetDirectoryName(outFile) + "\\report.txt");
                }

                // Check whether the document is valid or invalid.
                if (isValid == false) {
                    report.AddLog(input, "Translated DAISY file is not Valid: ", Report.ERROR_LEVEL);
                }

            } catch (UnauthorizedAccessException a) {
                xsd.Close();
                //dont have access permission
                String error = a.Message;
                report.AddLog(input, "Validation Error of converted DAISY File: \n" + error, Report.ERROR_LEVEL);
            } catch (Exception a) {
                xsd.Close();
                //and other things that could go wrong
                String error = a.Message;
                report.AddLog(input, "Validation Error of converted DAISY File: \n" + error, Report.ERROR_LEVEL);
            }
        }

        /* Function to Capture all the Validity Errors*/
        public void MyValidationEventHandler(object sender, ValidationEventArgs args) {
            isValid = false;
            String error = " Line Number : " + args.Exception.LineNumber + " and " +
             " Line Position : " + args.Exception.LinePosition + Environment.NewLine +
             " Message : " + args.Message + Environment.NewLine + " Reference Text :  " + errorText + Environment.NewLine;

            report.AddLog(input, "Validation Error of translated DAISY File: \n" + error, Report.ERROR_LEVEL);
        }

        #endregion

        /*Function which displays information about Commands to user*/
        private static void usage() {
            Console.WriteLine("Usage: DaisyConverterTest.exe /I PathOrFilename [/O PathOrFilename] [/BATCH-DOCX] [/REPORT Filename] [/TITLE] [/CREATOR] [/PUBLISHER][/UID] [/M] [/APAGE] [/CPAGE] [/STYLE]");
            Console.WriteLine("  Where options are:");
            Console.WriteLine("     /I PathOrFilename  Name of the file to transform (or input folder in case of batch conversion)");
            Console.WriteLine("     /O PathOrFilename  Name of the output file (or output folder)");
            Console.WriteLine("     /BATCH-DOCX        Do a batch conversion over every DOCX file in the input folder (note: existing files will be replaced)");
            Console.WriteLine("     /REPORT Filename   Name of the report file that must be generated (existing files will be replaced)");
            Console.WriteLine("     /TITLE      Data   Title of the output file that must be generated");
            Console.WriteLine("     /CREATOR    Data   Creator of the output file that must be generated");
            Console.WriteLine("     /PUBLISHER  Data   Publisher of the output file that must be generated");
            Console.WriteLine("     /UID        Data   Uid of the report output that must be generated");
            Console.WriteLine("     /M       Filename  To translate Multiple Documents");
            Console.WriteLine("     /APAGE   To Translate the current document with Automatic PageNumber Style");
            Console.WriteLine("     /CPAGE   To Translate the current document with Custom PageNumber Style");
            Console.WriteLine("     /STYLE   To Translate the current document with Character Styles");
            Console.WriteLine("     /PREPROCESS   To preprocess the document and get the parameteres and shapes exported as images");
        }


        #region Document Properties

        /// <summary>
        /// Function to get the Title of the Current Document
        /// </summary>
        /// <returns>Title</returns>
        public String DocPropTitle(String input) {
            int titleFlag = 0;
            String styleVal = "";
            string msgConcat = "";
            Package pack;
            pack = Package.Open(input, FileMode.Open, FileAccess.ReadWrite);
            String strTitle = pack.PackageProperties.Title;
            pack.Close();
            if (strTitle != "" && strTitle != null)
                return strTitle;
            else {
                pack = Package.Open(input, FileMode.Open, FileAccess.ReadWrite);
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    packRelationship = searchRelation;
                    break;
                }
                Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(mainPartxml.GetStream());
                NameTable nt = new NameTable();
                pack.Close();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                XmlNodeList getParagraph = doc.SelectNodes("//w:body/w:p/w:pPr/w:pStyle", nsManager);
                for (int j = 0; j < getParagraph.Count; j++) {
                    XmlAttributeCollection paraGraphAttribute = getParagraph[j].Attributes;
                    for (int i = 0; i < paraGraphAttribute.Count; i++) {
                        if (paraGraphAttribute[i].Name == "w:val") {
                            styleVal = paraGraphAttribute[i].Value;
                        }
                        if (styleVal != "" && styleVal == "Title") {
                            XmlNodeList getStyle = getParagraph[j].ParentNode.ParentNode.SelectNodes("w:r", nsManager);
                            if (getStyle != null) {
                                for (int k = 0; k < getStyle.Count; k++) {
                                    XmlNode getText = getStyle[k].SelectSingleNode("w:t", nsManager);
                                    msgConcat = msgConcat + " " + getText.InnerText;
                                }
                            }
                            titleFlag = 1;
                            break;
                        }
                        if (titleFlag == 1) {
                            break;
                        }
                    }
                    if (titleFlag == 1) {
                        break;
                    }
                }

                strTitle = msgConcat;

            }
            return strTitle;
        }

        #endregion

        #region Check for Input file/folder

        /*Function which validates the path of the input file*/
        private void CheckPaths() {
            if (this.input == null) {
                throw new DaisyCommandLineException("Input is missing");
            }
            if (this.batch == Direction.DocxToXml) {
                this.CheckBatch();
            } else if (this.batch == Direction.BatchDocx) {
                this.CheckBatch();
            } else if (this.batch == Direction.BatchOwnDocx) {
                this.CheckSingleFile();
            } else {
                this.CheckSingleFile();
            }
        }

        /*Function which validates the path of the input folder*/
        private void CheckBatch() {
            if (!Directory.Exists(this.input)) {
                throw new DaisyCommandLineException("Input folder does not exist");
            }
            if (File.Exists(this.output)) {
                throw new DaisyCommandLineException("Output must be a folder");
            }
            if (this.output == null || this.output.Length == 0) {
                // use input folder as output folder
                this.output = this.input;
            }
            if (!Directory.Exists(this.output)) {
                try {
                    Directory.CreateDirectory(this.output);
                } catch (Exception) {
                    throw new DaisyCommandLineException("Cannot create output folder");
                }
            }
        }

        /*Function which validates the path of the input file*/
        private void CheckSingleFile() {
            if (!File.Exists(this.input)) {
                throw new DaisyCommandLineException("Input file does not exist");
            }

            string extension = null;
            string inputExtension = Path.GetExtension(this.input).ToLowerInvariant();
            string outputExtension = "";
            if (this.output != null) {
                if (this.output.LastIndexOf(".").ToString() == "")
                    outputExtension = this.output.Substring(this.output.LastIndexOf("."));
                else
                    outputExtension = "";
            }
            switch (inputExtension) {

                case ".docx":
                case ".DOCX":
                    if (outputExtension.ToLower().Equals(".xml") || outputExtension.Equals("")) {
                        this.transformDirection = Direction.DocxToXml;
                        extension = ".xml";
                    } else {
                        throw new DaisyCommandLineException("Output file extension is invalid");
                    }
                    break;
                default:
                    throw new DaisyCommandLineException("Input file extension [" + inputExtension + "] is not supported.");
            }


            if (!File.Exists(this.output) && (this.output == null)) {
                string outputPath = this.output;
                if (outputPath == null) {
                    // we take input path
                    outputPath = Path.GetDirectoryName(this.input);
                }
                this.output = GenerateOutputName(outputPath, this.input, extension);
            }
        }

        #endregion

        #region Create Output Name

        /*Function to generate Name for output file*/
        private string GenerateOutputName(string rootPath, string input, string extension) {
            string rawFileName = Path.GetFileNameWithoutExtension(input);
            string output = Path.Combine(rootPath, rawFileName + extension);
            int num = 0;
            while (File.Exists(output)) {
                output = Path.Combine(rootPath, rawFileName + "_" + ++num + extension);
            }
            return output;
        }

        /*Function to generate Name for output file*/
        private string GenerateOutputName(string output, string extension) {
            string outputPath = output + extension;
            int num = 0;
            while (File.Exists(outputPath)) {
                outputPath = output + "_" + ++num + extension;
            }
            return outputPath;
        }

        #endregion



    }

}
