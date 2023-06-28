using System;
using System.Collections.Generic;
using System.IO;
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
            /// Possible type of note numbering outputed
            /// </summary>
            public enum Enum
            {
                /// <summary>
                /// Inline note in content (after the paragraph containing its first reference)
                /// </summary>
                Original,
                /// <summary>
                /// Put notes at the end of a level defined in settings
                /// </summary>
                Resize,
                /// <summary>
                /// Put the notes near the word pagebreak
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
        }

        #region Private fields with default values

        private string imgoption = "original";
        private string resampleValue = "96";
        private string characterStyle = "False";
        private string pagenumStyle = "Custom";
        private string footnotesLevel = "0"; // 0 mean current paragraphe, < 0 means parent level going upward, > 1 means absolute dtbook level
        private string footnotesPosition = "page"; // Should be inline, end, or page
        private string footnotesNumbering = "number"; // should be number, none, or word (to reuse word 
        private string footnotesStartValue = "1"; // number to be used
        private string footnotesNumberingPrefix = ""; // prefix to be added before the numbering
        private string footnotesNumberingSuffix = " "; // suffix to be added between the number and the text (default to a space to reproduce the plugin original behaviour)


        private string SettingsXml
        {
            get =>
                $"<Settings>" +
                    $"<PageNumbers  value=\"{pagenumStyle}\" />" +
                    $"<CharacterStyles value=\"{characterStyle}\" />" +
                    $"<ImageSizes value=\"{imgoption}\" samplingvalue=\"{resampleValue}\" />" +
                    $"<Footnotes level=\"{footnotesLevel}\" " +
                        $"position=\"{footnotesPosition}\" " +
                        $"numbering=\"{footnotesNumbering}\" " +
                        $"startValue=\"{footnotesStartValue}\" " +
                        $"numberPrefix=\"{footnotesNumberingPrefix}\" " +
                        $"numberSuffix=\"{footnotesNumberingSuffix}\" />" +
                $"</Settings>";

        }


        #endregion

        public void CreateDefaultSettings()
        {
            if (!Directory.Exists(ApplicationDataFolder))
            {
                Directory.CreateDirectory(ApplicationDataFolder);
            }
            using (StreamWriter writer = new StreamWriter(File.Create(ConverterSettingsFile)))
            {
                writer.Write(SettingsXml);
                writer.Flush();
            }
        }



        private ConverterSettings()
        {
            XmlDocument settingsDocument = new XmlDocument();
            if (!File.Exists(ConverterSettingsFile))
            {
                // Save the default settings
                save();
            }

            settingsDocument.Load(ConverterSettingsFile);
            XmlNode ImageSizesNode = settingsDocument.SelectSingleNode("//Settings/ImageSizes");
            if (ImageSizesNode != null)
            {
                imgoption = (ImageSizesNode.Attributes["value"]?.InnerXml) ?? imgoption;
                resampleValue = (ImageSizesNode.Attributes["samplingvalue"]?.InnerXml) ?? imgoption;
            }

            XmlNode CharacterStylesNode = settingsDocument.SelectSingleNode("//Settings/CharacterStyles");
            if (CharacterStylesNode != null)
            {
                characterStyle = (CharacterStylesNode.Attributes["value"]?.InnerXml) ?? characterStyle;
            }

            XmlNode PageNumbersNode = settingsDocument.SelectSingleNode("//Settings/PageNumbers");
            if (PageNumbersNode != null)
            {
                pagenumStyle = (PageNumbersNode.Attributes["value"]?.InnerXml) ?? pagenumStyle;
            }

            XmlNode FootnotesSettings = settingsDocument.SelectSingleNode("//Settings/Footnotes");
            if (FootnotesSettings != null)
            {
                footnotesLevel = (FootnotesSettings.Attributes["level"].InnerXml) ?? footnotesLevel;
                footnotesPosition = (FootnotesSettings.Attributes["position"]?.InnerXml) ?? footnotesPosition;
                footnotesNumbering = (FootnotesSettings.Attributes["numbering"]?.InnerXml) ?? footnotesNumbering;
                footnotesStartValue = (FootnotesSettings.Attributes["startValue"]?.InnerXml) ?? footnotesStartValue;
                footnotesNumberingPrefix = (FootnotesSettings.Attributes["numberPrefix"]?.InnerXml) ?? footnotesNumberingPrefix;
                footnotesNumberingSuffix = (FootnotesSettings.Attributes["numberSuffix"]?.InnerXml) ?? footnotesNumberingSuffix;
            }


        }

        /// <summary>
        /// Save the converter settings to an xml file on disk
        /// </summary>
        public void save()
        {
            if (!Directory.Exists(ApplicationDataFolder))
            {
                Directory.CreateDirectory(ApplicationDataFolder);
            }
            using (StreamWriter writer = new StreamWriter(File.Create(ConverterSettingsFile)))
            {
                writer.Write(SettingsXml);
                writer.Flush();
            }
        }



        public ImageOptionChoice.Enum ImageOption { get => ImageOptionChoice.Keys[imgoption]; set => imgoption = ImageOptionChoice.Values[value]; }

        public int ImageResamplingValue { get => int.Parse(resampleValue); set => resampleValue = value.ToString(); }

        public bool CharacterStyle { get => characterStyle != "False"; set => characterStyle = value ? "True" : "False"; }

        public string PagenumStyle { get => pagenumStyle; set => pagenumStyle = value; }

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
        public int FootnotesLevel { get => int.Parse(footnotesLevel); set => footnotesLevel = value.ToString(); }

        public FootnotesNumberingChoice.Enum FootnotesNumbering { get => FootnotesNumberingChoice.Keys[footnotesNumbering]; set => footnotesNumbering = FootnotesNumberingChoice.Values[value]; }

        public int FootnotesStartValue { get => int.Parse(footnotesStartValue); set => footnotesStartValue = value.ToString(); }

        public string FootnotesNumberingPrefix { get => footnotesNumberingPrefix; set => footnotesNumberingPrefix = value; }

        public string FootnotesNumberingSuffix { get => footnotesNumberingSuffix; set => footnotesNumberingSuffix = value; }


    }
}
