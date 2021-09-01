using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {
    /// <summary>
    /// Input parameters for a conversion.
    /// 
    /// This class also include the former TranslationParametersBuilder and PrepopulatedDaisyXml
    /// TODO : maybe this class shoudl include the script parser
    /// </summary>
    public class ConversionParameters
    {
        public ConverterSettings GlobalSettings = new ConverterSettings();

        /// <summary>
		/// Path for prepopulated_daisy xml file (conversion parameters cache file)
		/// </summary>
		private string ConversionParametersXmlPath {
            get { return ConverterHelper.AppDataSaveAsDAISYDirectory + "\\prepopulated_daisy.xml"; }
        }


        public string ControlName { get; set; }
		public string ScriptPath { get; set; }

        public string TempOutputFile { get; set; }
        

        // From the "TranslationParametersBuilder" and PrepopulatedDaisyXml class
        public string OutputPath { get; set; }
        public string PipelineOutput { get; set; }
        public string Title { get; set; }
        public string Creator { get; set; }
        public string Publisher { get; set; }
        public string UID { get; set; }
        
        public string Subject { get; set; }
        public string Version { get; set; }

        public FilenameValidator NameValidator { get; set; }

        public ScriptParser PostProcessSettings { get; set; } = null;

        /// <summary>
        /// Flag if changes should be tracked
        /// Possible values are Yes, No, or NoTrack
        /// </summary>
        public string TrackChanges { get; set; }

        /// <summary>
        /// Flag if sub documents should be parsed when found
        /// possible values are Yes, No or NoMasterFlag
        /// </summary>
        public string ParseSubDocuments { get; set; }

        /// <summary>
        /// Request DTBook XML validation
        /// </summary>
        public bool Validate { get; set; }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="wordVersion">Version of word use for the conversion</param>
        /// <param name="pipelineScript">Path of the pipeline script to use for postprocessing</param>
        /// <param name="filenameValidator">File name validator with regex matcher. If null, will default to dtbook XML filename pattern
        /// (see the one declared in ConverterHelper)</param>
        /// <param name="mainDocument">Document to retrieve conversion parameters from (Creator, Title and Publisher) </param>
        public ConversionParameters(string wordVersion = null, string pipelineScript = null, FilenameValidator filenameValidator = null, DocumentParameters mainDocument = null) {
            Version = wordVersion;
            ScriptPath = pipelineScript;
            if (pipelineScript != null) {
                PostProcessSettings = new ScriptParser(pipelineScript);
            }
            if(filenameValidator != null) {
                NameValidator = filenameValidator;
            } else {
                // Default to dtbook validator
                NameValidator = ConverterHelper.DTBookXMLFileNameFormat;
            }
            if (mainDocument != null) {
                this.usingMainDocument(mainDocument);
            } else if (File.Exists(ConversionParametersXmlPath)) { // Retrieve previous settings
                this.usingCachedSettings();
            }
        }

        public ConversionParameters usingMainDocument(DocumentParameters mainDocument) {
            Creator = PackageUtilities.DocPropCreator(mainDocument.CopyPath ?? mainDocument.InputPath);
            Title = PackageUtilities.DocPropTitle(mainDocument.CopyPath ?? mainDocument.InputPath);
            Publisher = PackageUtilities.DocPropPublish(mainDocument.CopyPath ?? mainDocument.InputPath);
            /*if(mainDocument.SubDocuments.Count > 0) {
                if(eventsHandler != null) {
                    ParseSubDocuments = eventsHandler.AskForTranslatingSubdocuments() ? "Yes" : "No";
                } else {
                    ParseSubDocuments = "No";
                }
            } else {
                ParseSubDocuments = "NoMasterSub";
            }*/
            return this;
        }

        public ConversionParameters usingCachedSettings() {
            XmlDocument document = new XmlDocument();
            document.Load(ConversionParametersXmlPath);
            Creator = document.FirstChild.ChildNodes[0].InnerText;
            Title = document.FirstChild.ChildNodes[1].InnerText;
            Publisher = document.FirstChild.ChildNodes[2].InnerText;
            return this;
        }




        /// <summary>
        /// Function to mimic the TranslationParametersBuilder with* construction
        /// 
        /// </summary>
        /// <param name="name">Name of the Class field to set</param>
        /// <param name="value">Object to assign to the field (this object will type casted to the targeted parameter type) </param>
        /// <returns>The converter itself</returns>
        public ConversionParameters withParameter(string name, object value) {
            switch (name) {
                case "ScriptPath":
                    ScriptPath = (string)value; break;
                case "OutputFile":
                    OutputPath = (string)value; break;
                case "Title":
                    Title = (string)value; break;
                case "Creator":
                    Creator = (string)value; break;
                case "Publisher":
                    Publisher = (string)value; break;
                case "UID":
                    UID = (string)value; break;
                case "Subject":
                    Subject = (string)value; break;
                case "Version":
                    Version = (string)value; break;
                case "PipelineOutput":
                    PipelineOutput = (string)value; break;
                case "PostProcessSetting":
                    PostProcessSettings = (ScriptParser)value; break;
                // Fields to be moved to document settings
                case "TrackChanges":
                    TrackChanges = (string)value; break;
                case "ParseSubDocuments":
                    ParseSubDocuments = (string)value; break;
                default:
                    break;
            }
            return this;
        }

        /// <summary>
        /// Get the conversion settings hashtable (to replace the TranslationParametersBuilder behavior)
        /// </summary>
        public Hashtable ConversionParametersHash {
            get {
                Hashtable parameters = new Hashtable();
                
                if (Title != null) parameters.Add("Title", Title);
                if (Creator != null) parameters.Add("Creator", Creator);
                if (Publisher != null) parameters.Add("Publisher", Publisher);
                if (UID != null) parameters.Add("UID", UID);
                if (Subject != null) parameters.Add("Subject", Subject);
                if (Version != null) parameters.Add("Version", Version);
                // TO BE CHANGED if the value changes in xslts
                if (OutputPath != null) parameters.Add("OutputFile", OutputPath);
                
                // This two should be moved to document parameters
                if (TrackChanges != null) parameters.Add("TRACK", TrackChanges);
                if (ParseSubDocuments != null) parameters.Add("MasterSub", ParseSubDocuments);

                // also retrieve global settings
                if (GlobalSettings.GetImageOption != " ") {
                    parameters.Add("ImageSizeOption", GlobalSettings.GetImageOption);
                    parameters.Add("DPI", GlobalSettings.GetResampleValue);
                }
                if (GlobalSettings.GetCharacterStyle != " ") {
                    parameters.Add("CharacterStyles", GlobalSettings.GetCharacterStyle);
                }
                if (GlobalSettings.GetPagenumStyle != " ") {
                    parameters.Add("Custom", GlobalSettings.GetPagenumStyle);
                }

                return parameters;
            }
        }

        /// <summary>
		/// Save current publisher/title/creator values to xml file.
		/// </summary>
		public void Save() {
            XmlDocument document = new XmlDocument();

            XmlElement elmtDaisy = document.CreateElement("Daisy");
            document.AppendChild(elmtDaisy);

            XmlElement elmtCreator, elmtTitle, elmtPublisher;

            elmtCreator = document.CreateElement("Creator");
            elmtDaisy.AppendChild(elmtCreator);
            elmtCreator.InnerText = Creator;

            elmtTitle = document.CreateElement("Title");
            elmtDaisy.AppendChild(elmtTitle);
            elmtTitle.InnerText = Title;

            elmtPublisher = document.CreateElement("Publisher");
            elmtDaisy.AppendChild(elmtPublisher);
            elmtPublisher.InnerText = Publisher;
            
            if (!System.IO.Directory.Exists(ConverterHelper.AppDataSaveAsDAISYDirectory))
                System.IO.Directory.CreateDirectory(ConverterHelper.AppDataSaveAsDAISYDirectory);

            document.Save(ConversionParametersXmlPath);
        }
    }
}