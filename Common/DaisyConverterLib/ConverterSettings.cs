using System;
using System.IO;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {
    /// <summary>
    /// Global settings of the converter, controled by "ConverterSettingsForm" class.<br/>
    /// Those settings are usually stored in a "DAISY_settingsVer21.xml" file in the application APPDATA directory
    /// </summary>
    public class ConverterSettings {
        #region Private and default fields
        private string imgoption;
        private int resampleValue = 96;
        private bool characterStyle = false;
        private string pagenumStyle;
        private int footnotesLevel = 0; // 0 mean current paragraphe, < 0 means parent level going upward, > 1 means absolute dtbook level
        private string footnotesPosition = "page" ; // Should be inline, end, after or page
        private string footnotesNumbering = "number"; // should be number or none (or empty, equals to none)
        private string footnotesStartValue = "1"; // can be a nummber to be used as starting number, or a character
        private string footnotesTextPrefix = " "; // prefix to be added between the numbering and the text (default to a simple space)

        private const string DefaultSettings = "" +
            "<Settings>" +
                "<PageNumbers  value=\"Custom\"/>" +
                "<CharacterStyles value=\"False\" />" +
                "<ImageSizes value=\"original\" samplingvalue=\"96\"/>" +
                "<Footnotes " +
                    "level=\"0\" " +
                    "position=\"page\" " +
                    "numbering=\"number\" " +
                    "startValue=\"1\" " +
                    "textPrefix=\"\"" +
            "/>" +
            "</Settings>";

        private const string SettingsFileName = "\\DAISY_settingsVer21.xml";
        #endregion

        public static void CreateDefaultSettings(string settingsFolder) {
            using (StreamWriter writer = new StreamWriter(File.Create(settingsFolder + SettingsFileName))) {
                writer.Write(DefaultSettings);
                writer.Flush();
            }
        }

        public ConverterSettings() {
            //String xmlfile_path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            String xmlfile_path = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY");
            XmlDocument translationxml = new XmlDocument();
            if (!File.Exists(xmlfile_path + SettingsFileName)) {
                CreateDefaultSettings(xmlfile_path);
            }

            translationxml.Load(xmlfile_path + SettingsFileName);
            XmlNode imgnodeOption = translationxml.SelectSingleNode("//Settings/ImageSizes[@value]");
            imgoption = imgnodeOption.Attributes[0].InnerXml;
            XmlNode imgResampleValue = translationxml.SelectSingleNode("//Settings/ImageSizes");
            resampleValue = Convert.ToInt32(imgResampleValue.Attributes[1].Value);
            XmlNode charstyleOption = translationxml.SelectSingleNode("//Settings/CharacterStyles[@value]");
            characterStyle = charstyleOption.Attributes[0].InnerXml == "True" ? true : false;
            XmlNode pagenumOption = translationxml.SelectSingleNode("//Settings/PageNumbers[@value]");
            pagenumStyle = pagenumOption.Attributes[0].InnerXml;

            XmlNode FootnotesSettings = translationxml.SelectSingleNode("//Settings/Footnotes");
            footnotesLevel = Convert.ToInt32(
                (FootnotesSettings != null && FootnotesSettings.Attributes["level"] != null) 
                ? FootnotesSettings.Attributes["level"].InnerXml
                : "0"
            );

            footnotesPosition = (FootnotesSettings != null && FootnotesSettings.Attributes["position"] != null)
                ? FootnotesSettings.Attributes["position"].InnerXml
                : "page";

        }

        public string GetImageOption => imgoption;

        public int GetResampleValue => resampleValue;

        public bool GetCharacterStyle => characterStyle;

        public string GetPagenumStyle => pagenumStyle;

        /// <summary>
        /// Position of the notes relatively to the selected level <br/>
        /// - <b>(Default)</b> page override the level selection and insert the note when the next pagebreak is found <br/>
        /// - inline means it will placed within the selected level after the child containing the noteref <br/>
        /// - end means it will placed before the end tag of the selected level <br/>
        /// - after means it will placed right after the closing tag of the selected level <br/>
        /// </summary>
        public string FootnotesPosition { get => footnotesPosition; set => footnotesPosition = value; }

        /// <summary>
        /// Level where the footnote must be placed relatively to its first reference call.<br/>
        /// - <b>(Default)</b> a 0 value means that it must be placed relative to the current bloc (likely a paragraph, a list or a title) <br/>
        /// - a negative value (-N) means it will placed relative to the Nth parent of the current bloc,<br/>
        /// - a positive value (N) means it will placed relatively to the current absolute N level
        /// </summary>
        public int FootnotesLevel { get => footnotesLevel; set => footnotesLevel = value; }
    }
}
