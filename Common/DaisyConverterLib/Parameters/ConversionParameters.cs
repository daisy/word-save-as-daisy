using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

using Newtonsoft.Json;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion
{
    /// <summary>
    /// Input parameters for a conversion.
    /// 
    /// This class also include the former TranslationParametersBuilder and PrepopulatedDaisyXml
    /// TODO : maybe this class shoudl include the script parser
    /// </summary>
    public class ConversionParameters
    {
        public ConverterSettings GlobalSettings = Instance;

        /// <summary>
		/// Path for prepopulated_daisy xml file (conversion parameters cache file)
		/// </summary>
		private string ConversionParametersXmlPath
        {
            get { return ConverterHelper.AppDataSaveAsDAISYDirectory + "\\prepopulated_daisy.xml"; }
        }


        public string ControlName { get; set; }


        // From the "TranslationParametersBuilder" and PrepopulatedDaisyXml class
        public string OutputPath { get; set; }
        public string PipelineOutput { get; set; }
        public string Title { get; set; }
        public string Creator { get; set; }
        public string Publisher { get; set; }
        public string UID { get; set; }

        public string Subject { get; set; }
        public string Version { get; set; }

        public string Language { get; set; } = "";

        public StringValidator NameValidator { get; set; }

        public Script PostProcessor { get; set; } = null;

        /// <summary>
        /// Flag if changes should be tracked
        /// Possible values are Yes, No, or NoTrack
        /// </summary>
        public string TrackChanges { get; set; }

        /// <summary>
        /// Flag if sub documents should be parsed when found
        /// possible values are Yes, No or NoMasterFlag
        /// </summary>
        public bool ParseSubDocuments { get; set; } = false;

        /// <summary>
        /// Request DTBook XML validation
        /// </summary>
        public bool Validate { get; set; }


        /// <summary>
        /// Keep word visible during conversion
        /// </summary>
        public bool Visible { get; set; } = true;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wordVersion">Version of word use for the conversion</param>
        /// <param name="pipelineScript">Path of the pipeline script to use for postprocessing</param>
        /// <param name="filenameValidator">File name validator with regex matcher. If null, will default to dtbook XML filename pattern
        /// (see the one declared in ConverterHelper)</param>
        /// <param name="mainDocument">Document to retrieve conversion parameters from (Creator, Title and Publisher) </param>
        public ConversionParameters(string wordVersion = null, Script pipelineScript = null, StringValidator filenameValidator = null, DocumentParameters mainDocument = null)
        {
            Version = wordVersion;
            //ScriptPath = pipelineScript;
            if (pipelineScript != null)
            {
                PostProcessor = pipelineScript;
            }
            if (filenameValidator != null)
            {
                NameValidator = filenameValidator;
            }
            else
            {
                // Default to dtbook validator
                NameValidator = StringValidator.DTBookXMLFileNameFormat;
            }
            if (mainDocument != null)
            {
                this.usingMainDocument(mainDocument);
            }
            else if (File.Exists(ConversionParametersXmlPath))
            { // Retrieve previous settings
                this.usingCachedSettings();
            }
        }

        public ConversionParameters usingMainDocument(DocumentParameters mainDocument)
        {
            // Analyze the copy to retrieve document properties
            mainDocument.updatePropertiesFromCopy();
            Creator = mainDocument.Creator;
            Title = mainDocument.Title;
            Publisher = mainDocument.Publisher;
            Language = mainDocument.Languages[0];
            return this;
        }

        public ConversionParameters usingCachedSettings()
        {
            XmlDocument document = new XmlDocument();
            try {
                document.Load(ConversionParametersXmlPath);
                Creator = document.FirstChild.ChildNodes[0].InnerText;
                Title = document.FirstChild.ChildNodes[1].InnerText;
                Publisher = document.FirstChild.ChildNodes[2].InnerText;
            } catch (Exception e) {
               AddinLogger.Warning($"Could not load cache from {ConversionParametersXmlPath}", e);
            }


            return this;
        }




        /// <summary>
        /// Function to mimic the TranslationParametersBuilder with* construction
        /// 
        /// </summary>
        /// <param name="name">Name of the Class field to set</param>
        /// <param name="value">Object to assign to the field (this object will type casted to the targeted parameter type) </param>
        /// <returns>The converter itself</returns>
        public ConversionParameters withParameter(string name, object value)
        {
            switch (name)
            {
                case "Language":
                    Language = (string)value; break;
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
                case "PostProcessor":
                    PostProcessor = (Script)value; break;
                // Fields to be moved to document settings
                case "TrackChanges":
                    TrackChanges = (string)value; break;
                case "ParseSubDocuments":
                    ParseSubDocuments = (bool)value; break;
                case "Visible":
                    Visible = (bool)value; break;
                default:
                    break;
            }
            return this;
        }

        /// <summary>
        /// Get the conversion settings hashtable (to replace the TranslationParametersBuilder behavior)
        /// </summary>
        public Hashtable ParametersHash
        {
            get
            {
                Hashtable parameters = new Hashtable();

                if (Title != null) parameters.Add("Title", Title);
                if (Creator != null) parameters.Add("Creator", Creator);
                if (Publisher != null) parameters.Add("Publisher", Publisher);
                if (UID != null) parameters.Add("UID", UID);
                if (Subject != null) parameters.Add("Subject", Subject);
                if (Version != null) parameters.Add("Version", Version);
                // TO BE CHANGED if the value changes in xslts
                if (OutputPath != null) parameters.Add("OutputFile", OutputPath);

                // This two fields are stored in document parameters,
                // but we let the conversion settings capable of overriding them just in case
                if (TrackChanges != null) parameters.Add("TRACK", TrackChanges);
                parameters.Add("MasterSub", ParseSubDocuments);

                // also retrieve global settings
                parameters.Add("ImageSizeOption", GlobalSettings.ImageOption);
                parameters.Add("DPI", GlobalSettings.ImageResamplingValue);

                parameters.Add("CharacterStyles", GlobalSettings.CharacterStyle);

                if (GlobalSettings.PagenumStyle != " ")
                {
                    parameters.Add("Custom", GlobalSettings.PagenumStyle);
                }
                // 20220402 : adding footnotes positioning settings
                // might be "page", "inline", "end" or "after"
                parameters.Add("FootnotesPosition", FootnotesPositionChoice.Values[GlobalSettings.FootnotesPosition]);

                // value between -5 and 6, with negative value meaning parents level going upwards,
                // 0 meaning in current level,
                // positive value meaning absolute level value where to put the notes
                parameters.Add("FootnotesLevel", GlobalSettings.FootnotesLevel);

                // 20230113 : adding footnotes numbering customization
                parameters.Add("FootnotesNumbering", FootnotesNumberingChoice.Values[GlobalSettings.FootnotesNumbering]);
                parameters.Add("FootnotesStartValue", GlobalSettings.FootnotesStartValue);
                parameters.Add("FootnotesNumberingPrefix", GlobalSettings.FootnotesNumberingPrefix);
                parameters.Add("FootnotesNumberingSuffix", GlobalSettings.FootnotesNumberingSuffix);

                // 20240220 : adding selected language
                parameters.Add("Language", Language);

                return parameters;
            }
        }

        /// <summary>
		/// Save current publisher/title/creator values to xml file.
		/// </summary>
		public void Save()
        {
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

        public string serialize()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}