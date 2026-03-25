using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    /// <summary>
    /// Settable Properties modified by the user (like API keys for external services) <br/>
    /// Those properties are to be sent to webservice or jni wrapper to enable other features in the used engine (like TTS and OCR). <br/>
    /// </summary>
    public class PipelineUserProperties
    {

        public static string PipelineUserPropertiesFile = Path.Combine(ConverterSettings.ApplicationDataFolder, "engine.user.props");

        #region Singleton definition
        private static readonly Lazy<PipelineUserProperties> lazy = new Lazy<PipelineUserProperties>(() => new PipelineUserProperties());

        public static PipelineUserProperties Instance => lazy.Value;


        public List<EngineProperty> Items { get; private set; } = new List<EngineProperty>();

        private PipelineUserProperties() {
            if(File.Exists(PipelineUserPropertiesFile))
            {
                try
                {
                    System.Xml.XmlDocument propsDoc = new System.Xml.XmlDocument();
                    propsDoc.Load(PipelineUserPropertiesFile);
                    XmlElement propertiesNode = propsDoc.GetElementsByTagName("properties")[0] as XmlElement;
                    Items = EngineProperty.ListFromXml(propertiesNode);
                }
                catch (Exception ex)
                {
                    AddinLogger.Error("An error occured while loading user engine properties file", ex);
                }
            }
        }

        public List<EngineProperty> mergeInto(List<EngineProperty> items)
        {
            foreach (var item in items)
            {
                var userItem = Items.FirstOrDefault(i => i.Name == item.Name);
                if (userItem != null)
                {
                    item.Value = userItem.Value;
                }
            }
            return items;
        }

        public void ReplaceBy(List<EngineProperty> properties)
        {
            Items = properties;
        }

        public void UpdateOrAddRange(List<EngineProperty> properties)
        {
            foreach (var property in properties)
            {
                var existingItem = Items.FirstOrDefault(i => i.Name == property.Name);
                if (existingItem != null)
                {
                    existingItem.Value = property.Value;
                }
                else
                {
                    Items.Add(property);
                }
            }
        }

        public void Save()
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                XmlElement root = doc.CreateElement("properties");
                doc.AppendChild(root);
                foreach (var item in Items)
                {
                    XmlElement itemNode = item.ToXml(doc);
                    root.AppendChild(itemNode);
                }
                doc.Save(PipelineUserPropertiesFile);
            }
            catch (Exception ex)
            {
                AddinLogger.Error("An error occured while saving user engine properties file", ex);
            }
        }



        #endregion
    }
}
