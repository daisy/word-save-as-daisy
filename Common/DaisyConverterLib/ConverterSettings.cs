using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion
{
    /// <summary>
    /// Global settings of the converter, controled by "ConverterSettingsForm" class.<br/>
    /// Those settings are usually stored in a "DAISY_settingsVer21.xml" file in the application APPDATA directory
    /// </summary>
    public sealed class ConverterSettings
    {

        public static string ApplicationDataFolder = Path.GetFullPath(
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "SaveAsDAISY"
            )
        );

        public static string DefaultResultsFolder = Path.GetFullPath(
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Documents",
                "SaveAsDAISY Results"
            )
        );

        private static string ConverterSettingsFile = Path.Combine(ApplicationDataFolder, "DAISY_settingsVer21.xml");

        #region Singleton definition
        private static readonly Lazy<ConverterSettings> lazy = new Lazy<ConverterSettings>(() => new ConverterSettings());

        public static ConverterSettings Instance => lazy.Value;
        #endregion
        /// <summary>
        /// @see FootnotesNumberingChoice.Enum
        /// </summary>
        public static class ImageOptionChoice
        {
            /// <summary>
            /// Image resizing
            /// </summary>
            public enum Enum
            {
                /// <summary>
                /// Keep original image
                /// </summary>
                Original,
                /// <summary>
                /// Resize image
                /// </summary>
                Resize,
                /// <summary>
                /// Resample the image
                /// </summary>
                Resample
            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.Original, "original" },
                { Enum.Resize, "resize" },
                { Enum.Resample, "resample" },
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                { "original", Enum.Original },
                { "resize", Enum.Resize },
                { "resample", Enum.Resample},
            };

            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.ImageOption]
                );
            }
        }

        public static class ImageResamplingChoice
        {
        /// <summary>
            /// Possible type of note numbering outputed
            /// </summary>
            public enum Enum
            {
                dpi_72,
                dpi_96,
                dpi_120,
                dpi_150,
                dpi_300,

            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.dpi_72, "72" },
                { Enum.dpi_96, "96" },
                { Enum.dpi_120, "120" },
                { Enum.dpi_150, "150" },
                { Enum.dpi_300, "300" },
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                { "72" , Enum.dpi_72 },
                { "96" , Enum.dpi_96 },
                { "120" , Enum.dpi_120 },
                { "150" , Enum.dpi_150 },
                { "300" , Enum.dpi_300 },
            };
            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.ImageResamplingValue]
                );
            }
        }

        public static class FootnotesLevelChoice
        {
            /// <summary>
            /// Possible type of note numbering outputed
            /// </summary>
            public enum Enum
            {
                Inlined,
                Level_1,
                Level_2,
                Level_3,
                Level_4,
                Level_5,
                Level_6,
            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.Inlined, "0" },
                { Enum.Level_1, "1" },
                { Enum.Level_2, "2" },
                { Enum.Level_3, "3" },
                { Enum.Level_4, "4" },
                { Enum.Level_5, "5" },
                { Enum.Level_6, "6" }
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                {"0" , Enum.Inlined},
                {"1" , Enum.Level_1},
                {"2" , Enum.Level_2},
                {"3" , Enum.Level_3},
                {"4" , Enum.Level_4},
                {"5" , Enum.Level_5},
                {"6" , Enum.Level_6}
            };
            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.FootnotesLevel]
                );
            }
        }

        /// <summary>
        /// @see FootnotesNumberingChoice.Enum
        /// </summary>
        public static class FootnotesPositionChoice
        {
            /// <summary>
            /// Possible type of note numbering outputed
            /// </summary>
            public enum Enum
            {
                /// <summary>
                /// Inline note in content (after the paragraph containing its first reference)
                /// </summary>
                Inline,
                /// <summary>
                /// Put notes at the end of a level defined in settings
                /// </summary>
                End,
                /// <summary>
                /// Put the notes near the word pagebreak
                /// </summary>
                Page
            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.Inline, "inline" },
                { Enum.End, "end" },
                { Enum.Page, "page" },
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                { "inline", Enum.Inline },
                { "end", Enum.End },
                { "page", Enum.Page},
            };

            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.FootnotesPosition]
                );
            }
        }

        /// <summary>
        /// @see FootnotesNumberingChoice.Enum
        /// </summary>
        public static class FootnotesNumberingChoice
        {
            /// <summary>
            /// Possible type of note numbering outputed
            /// </summary>
            public enum Enum
            {
                /// <summary>
                /// Use original word numbering
                /// </summary>
                Word,
                /// <summary>
                /// Use custom numbering, starting from the settings start value
                /// </summary>
                Number,
                /// <summary>
                /// Disable note numbering in output
                /// </summary>
                None
            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.Word, "word" },
                { Enum.Number, "number" },
                { Enum.None, "none" },
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                { "word", Enum.Word },
                { "number", Enum.Number },
                { "none", Enum.None},
            };

            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.FootnotesNumbering]
                );
            }
        }

        public static class PageNumberingChoice
        {
            /// <summary>
            /// Possible type of page numbering computation
            /// </summary>
            public enum Enum
            {
                /// <summary>
                /// Use numbers tagged with style PagenumberDAISY to insert page numbers in content
                /// </summary>
                Custom,
                /// <summary>
                /// Use word page break to compute and insert page numbers in content
                /// </summary>
                Automatic
            }
            public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
            {
                { Enum.Custom, "custom" },
                { Enum.Automatic, "automatic" },
            };
            public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
            {
                { "custom", Enum.Custom },
                { "automatic", Enum.Automatic },
            };
            public static EnumData DataType()
            {
                return new EnumData(
                    Values.ToDictionary(kvp => kvp.Key.ToString(), kvp => (object)kvp.Value),
                    Values[Instance.PagenumStyle] // Default value is now stored here
                );
            }
        }

        #region Private fields with default values

        private string imgoption = ImageOptionChoice.Values[ImageOptionChoice.Enum.Original];
        private string resampleValue = "96";
        private string characterStyle = "False";
        private string pagenumStyle = PageNumberingChoice.Values[PageNumberingChoice.Enum.Custom];
        private string footnotesLevel = "0"; // 0 mean current paragraphe level, < 0 means parent level going upward, > 1 means absolute dtbook level
        private string footnotesPosition = FootnotesPositionChoice.Values[FootnotesPositionChoice.Enum.Inline]; // Should be inline, end, or page
        private string footnotesNumbering = FootnotesNumberingChoice.Values[FootnotesNumberingChoice.Enum.None]; // should be number, none, or word
        private string footnotesStartValue = "1"; // number to be used
        private string footnotesNumberingPrefix = ""; // prefix to be added before the numbering
        private string footnotesNumberingSuffix = ""; // suffix to be added between the number and the text
        //private string azureSpeechRegion = ""; // region defined in the Azure console for the speech service
        //private string azureSpeechKey = ""; // one of the two keys provided to connect to to Azure speech synthesis service
        private string ttsConfigFile = ""; // A tts config file to use for speech synthesis with pipeline 2
        private string dontNotifySponsorship = ""; // notify the user about sponsorship
        private bool useDAISYPipelineApp = true; // use the pipeline app to run the conversion instead of the embedded engine
        private string resultsFolder = DefaultResultsFolder; // Default results folder
        /// <summary>
        /// Get the current settings as XML string
        /// </summary>
        /// <param name="withPrivateData">if set to true, settings that are private to users 
        /// (like tts keys) will also be included</param>
        /// <returns></returns>
        public string AsXML(bool withPrivateData = false)
        {
            return $"<Settings>" +
                $"\r\n\t<PageNumbers  value=\"{pagenumStyle}\" />" +
                $"\r\n\t<CharacterStyles value=\"{characterStyle}\" />" +
                $"\r\n\t<ImageSizes value=\"{imgoption}\" samplingvalue=\"{resampleValue}\" />" +
                $"\r\n\t<Footnotes level=\"{footnotesLevel}\" " +
                $"\r\n\t\tposition=\"{footnotesPosition}\" " +
                $"\r\n\t\tnumbering=\"{footnotesNumbering}\" " +
                $"\r\n\t\tstartValue=\"{footnotesStartValue}\" " +
                $"\r\n\t\tnumberPrefix=\"{footnotesNumberingPrefix}\" " +
                $"\r\n\t\tnumberSuffix=\"{footnotesNumberingSuffix}\" />" +
                $"\r\n\t<TTSConfig file=\"{ttsConfigFile}\" />" +
                //(withPrivateData
                //    ? $"\r\n\t<Azure region=\"{azureSpeechRegion}\" " +
                //      $"\r\n\t\tkey=\"{azureSpeechKey}\" />"
                //    : ""
                //) +
                $"\r\n\t<DontNotifySponsorship value=\"{dontNotifySponsorship}\" />" +
                $"\r\n\t<UsePipelineApp value=\"{useDAISYPipelineApp.ToString().ToLower()}\" />" +
                $"\r\n\t<ResultsFolder value=\"{resultsFolder}\" />" +
                $"\r\n</Settings>";
        }


        #endregion

        public void CreateDefaultSettings()
        {
            if (!Directory.Exists(ApplicationDataFolder)) {
                Directory.CreateDirectory(ApplicationDataFolder);
            }
            using (StreamWriter writer = new StreamWriter(File.Create(ConverterSettingsFile))) {
                writer.Write(AsXML());
                writer.Flush();
            }
        }



        private ConverterSettings()
        {
            XmlDocument settingsDocument = new XmlDocument();
            if (!File.Exists(ConverterSettingsFile)) {
                // Save the default settings
                Save();
            }

            settingsDocument.Load(ConverterSettingsFile);
            XmlNode ImageSizesNode = settingsDocument.SelectSingleNode("//Settings/ImageSizes");
            if (ImageSizesNode != null) {
                imgoption = (ImageSizesNode.Attributes["value"]?.InnerXml) ?? imgoption;
                resampleValue = (ImageSizesNode.Attributes["samplingvalue"]?.InnerXml) ?? imgoption;
            }

            XmlNode CharacterStylesNode = settingsDocument.SelectSingleNode("//Settings/CharacterStyles");
            if (CharacterStylesNode != null) {
                characterStyle = (CharacterStylesNode.Attributes["value"]?.InnerXml) ?? characterStyle;
            }

            XmlNode PageNumbersNode = settingsDocument.SelectSingleNode("//Settings/PageNumbers");
            if (PageNumbersNode != null) {
                pagenumStyle = ((PageNumbersNode.Attributes["value"]?.InnerXml) ?? pagenumStyle).ToLowerInvariant();
            }

            XmlNode FootnotesSettings = settingsDocument.SelectSingleNode("//Settings/Footnotes");
            if (FootnotesSettings != null) {
                footnotesLevel = (FootnotesSettings.Attributes["level"].InnerXml) ?? footnotesLevel;
                footnotesPosition = (FootnotesSettings.Attributes["position"]?.InnerXml) ?? footnotesPosition;
                footnotesNumbering = (FootnotesSettings.Attributes["numbering"]?.InnerXml) ?? footnotesNumbering;
                footnotesStartValue = (FootnotesSettings.Attributes["startValue"]?.InnerXml) ?? footnotesStartValue;
                footnotesNumberingPrefix = (FootnotesSettings.Attributes["numberPrefix"]?.InnerXml) ?? footnotesNumberingPrefix;
                footnotesNumberingSuffix = (FootnotesSettings.Attributes["numberSuffix"]?.InnerXml) ?? footnotesNumberingSuffix;
            }

            //XmlNode AzureSettings = settingsDocument.SelectSingleNode("//Settings/Azure");
            //if (AzureSettings != null) {
            //    azureSpeechRegion = (AzureSettings.Attributes["region"].InnerXml) ?? azureSpeechRegion;
            //    azureSpeechKey = (AzureSettings.Attributes["key"].InnerXml) ?? azureSpeechKey;
            //}
            XmlNode TTSConfigSettings = settingsDocument.SelectSingleNode("//Settings/TTSConfig");
            if (TTSConfigSettings != null) {
                ttsConfigFile = (TTSConfigSettings.Attributes["file"].InnerXml) ?? ttsConfigFile;
            }

            XmlNode DontNotifySponsorshipSettings = settingsDocument.SelectSingleNode("//Settings/DontNotifySponsorship");
            if (DontNotifySponsorshipSettings != null) {
                dontNotifySponsorship = (DontNotifySponsorshipSettings.Attributes["value"].InnerXml) ?? dontNotifySponsorship;
            }
            XmlNode UsePipelineAppSettings = settingsDocument.SelectSingleNode("//Settings/UsePipelineApp");
            if (UsePipelineAppSettings != null) {
                var nodeValue = UsePipelineAppSettings.Attributes["value"]?.InnerXml;
                useDAISYPipelineApp = ConverterHelper.PipelineAppIsInstalled() && (nodeValue == null ? useDAISYPipelineApp : nodeValue == "true");
            }


        }

        /// <summary>
        /// Save the converter settings to an xml file on disk
        /// </summary>
        public void Save()
        {
            if (!Directory.Exists(ApplicationDataFolder)) {
                Directory.CreateDirectory(ApplicationDataFolder);
            }
            using (StreamWriter writer = new StreamWriter(File.Create(ConverterSettingsFile))) {
                writer.Write(AsXML(true));
                writer.Flush();
            }
            // NP 2025 10 07 : launch the adequate runner based on settings
            // Launch the pipeline in the background to start conversions asap
            if (UseDAISYPipelineApp) {
                try {
                    AppRunner.GetInstance();
                }
                catch (Exception e) {
                    AddinLogger.Error(e);
                }
            } else {
                try {
                    JNIRunner.GetInstance();
                }
                catch (Exception e) {
                    AddinLogger.Error(e);
                }
            }
        }



        public ImageOptionChoice.Enum ImageOption { get => ImageOptionChoice.Keys[imgoption]; set => imgoption = ImageOptionChoice.Values[value]; }

        public ImageResamplingChoice.Enum ImageResamplingValue { get => ImageResamplingChoice.Keys[resampleValue]; set => resampleValue = ImageResamplingChoice.Values[value]; }

        public bool CharacterStyle { get => characterStyle != "False"; set => characterStyle = value ? "True" : "False"; }

        public PageNumberingChoice.Enum PagenumStyle { get => PageNumberingChoice.Keys[pagenumStyle]; set => pagenumStyle = PageNumberingChoice.Values[value]; }

        /// <summary>
        /// Position of the notes relatively to the selected level <br/>
        /// - <b>(Default)</b> page override the level selection and insert the note when the next pagebreak is found <br/>
        /// - inline means it will placed within the selected level after the child containing the noteref <br/>
        /// - end means it will placed before the end tag of the selected level <br/>
        /// - after means it will placed right after the closing tag of the selected level <br/>
        /// </summary>
        public FootnotesPositionChoice.Enum FootnotesPosition { get => FootnotesPositionChoice.Keys[footnotesPosition]; set => footnotesPosition = FootnotesPositionChoice.Values[value]; }

        /// <summary>
        /// Level where the footnote must be placed relatively to its first reference call.<br/>
        /// - <b>(Default)</b> a 0 value means that it must be placed relative to the current bloc (likely a paragraph, a list or a title) <br/>
        /// - a negative value (-N) means it will placed relative to the Nth parent of the current bloc,<br/>
        /// - a positive value (N) means it will placed relatively to the current absolute N level
        /// </summary>
        public FootnotesLevelChoice.Enum FootnotesLevel { get => FootnotesLevelChoice.Keys[footnotesLevel]; set => footnotesLevel = FootnotesLevelChoice.Values[value]; }

        public FootnotesNumberingChoice.Enum FootnotesNumbering { get => FootnotesNumberingChoice.Keys[footnotesNumbering]; set => footnotesNumbering = FootnotesNumberingChoice.Values[value]; }

        public int FootnotesStartValue { get => int.Parse(footnotesStartValue); set => footnotesStartValue = value.ToString(); }

        public string FootnotesNumberingPrefix { get => footnotesNumberingPrefix; set => footnotesNumberingPrefix = value; }

        public string FootnotesNumberingSuffix { get => footnotesNumberingSuffix; set => footnotesNumberingSuffix = value; }

        //public string AzureSpeechRegion { get => azureSpeechRegion; set => azureSpeechRegion = value; }

        //public string AzureSpeechKey { get => azureSpeechKey; set => azureSpeechKey = value; }

        public string TTSConfigFile { get => ttsConfigFile; set => ttsConfigFile = value; }

        public bool DontNotifySponsorship {
            get
            {
                try {
                    return dontNotifySponsorship.Length > 0 && bool.Parse(dontNotifySponsorship);
                }
                catch (Exception) {
                    return false;
                }
            }
            set => dontNotifySponsorship = value.ToString();
        }

        public bool UseDAISYPipelineApp {
            get => useDAISYPipelineApp;
            set => useDAISYPipelineApp = value;
        }

        public string ResultsFolder {
            get => resultsFolder;
            set {
                resultsFolder = value ?? DefaultResultsFolder; // If null, use the default results folder
                if (Directory.Exists(value)) {
                    resultsFolder = value;
                } else {
                    resultsFolder = DefaultResultsFolder;
                    //throw new DirectoryNotFoundException($"The results folder '{value}' does not exist.");
                }
            }
        }
    }
}
