using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Text;
using System.Xml.Xsl;
using System.Windows.Forms;
using System.Drawing;
using System.Resources;
using System.Xml.XPath;
using System.Reflection;
using System.Collections;
using System.IO.Packaging;
using System.ComponentModel;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Runtime.InteropServices;


namespace Daisy.DaisyConverter.DaisyConverterLib
{
    class DaisyTranslationSettings
    {
        String imgoption;
        String resampleValue;
        String characterStyle;
        String pagenumStyle;

    	private const string DefaultSettings = "<Settings><PageNumbers  value=\"Custom\"/><CharacterStyles value=\"False\" /><ImageSizes value=\"original\" samplingvalue=\"96\"/></Settings>";

    	private const string SettingsFileName = "\\DAISY_settingsVer21.xml";

    	public static void CreateDefaultSettings(string settingsFolder)
		{
			using (StreamWriter writer = new StreamWriter(File.Create(settingsFolder + SettingsFileName)))
			{
				writer.Write(DefaultSettings);
				writer.Flush();
			}
		}

        public DaisyTranslationSettings()
        {
            //String xmlfile_path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            String xmlfile_path = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY");
            XmlDocument translationxml = new XmlDocument();
            if (!File.Exists(xmlfile_path + SettingsFileName))
            {
            	CreateDefaultSettings(xmlfile_path);
            }

			translationxml.Load(xmlfile_path + SettingsFileName);
            XmlNode imgnodeOption = translationxml.SelectSingleNode("//Settings/ImageSizes[@value]");
             imgoption = imgnodeOption.Attributes[0].InnerXml;
             XmlNode imgResampleValue = translationxml.SelectSingleNode("//Settings/ImageSizes");
             resampleValue=imgResampleValue.Attributes[1].Value;
             XmlNode charstyleOption = translationxml.SelectSingleNode("//Settings/CharacterStyles[@value]");
             characterStyle = charstyleOption.Attributes[0].InnerXml;
             XmlNode pagenumOption = translationxml.SelectSingleNode("//Settings/PageNumbers[@value]");
             pagenumStyle = pagenumOption.Attributes[0].InnerXml;

        }

        public String GetImageOption
        {
            get
            {
                return imgoption;
            }
        }

        public String GetResampleValue
        {
            get
            {
                return resampleValue;
            }
        }

        public String GetCharacterStyle
        {
            get
            {
                return characterStyle;
            }
        }
        public String GetPagenumStyle
        {
            get
            {
                return pagenumStyle;
            }
        }

    }
}
