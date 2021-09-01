using System;
using System.IO;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {
    /// <summary>
    /// Global settings of the converter, controled by "ConverterSettingsForm" class.<br/>
    /// Those settings are usually stored in a "DAISY_settingsVer21.xml" file in the application APPDATA directory
    /// </summary>
    public class ConverterSettings {
        string imgoption;
        string resampleValue;
        string characterStyle;
        string pagenumStyle;

        private const string DefaultSettings = "<Settings><PageNumbers  value=\"Custom\"/><CharacterStyles value=\"False\" /><ImageSizes value=\"original\" samplingvalue=\"96\"/></Settings>";

        private const string SettingsFileName = "\\DAISY_settingsVer21.xml";

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
            resampleValue = imgResampleValue.Attributes[1].Value;
            XmlNode charstyleOption = translationxml.SelectSingleNode("//Settings/CharacterStyles[@value]");
            characterStyle = charstyleOption.Attributes[0].InnerXml;
            XmlNode pagenumOption = translationxml.SelectSingleNode("//Settings/PageNumbers[@value]");
            pagenumStyle = pagenumOption.Attributes[0].InnerXml;

        }

        public string GetImageOption {
            get {
                return imgoption;
            }
        }

        public string GetResampleValue {
            get {
                return resampleValue;
            }
        }

        public string GetCharacterStyle {
            get {
                return characterStyle;
            }
        }
        public string GetPagenumStyle {
            get {
                return pagenumStyle;
            }
        }

    }
}
