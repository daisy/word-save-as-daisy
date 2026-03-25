using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion.Pipeline
{
    public sealed class EmbeddedEngineProperties
    {
        private static readonly Lazy<EmbeddedEngineProperties> lazy
            = new Lazy<EmbeddedEngineProperties>(() => new EmbeddedEngineProperties());

        public static EmbeddedEngineProperties Instance => lazy.Value;

        
        public List<EngineProperty> Items { get; private set; }

        private EmbeddedEngineProperties()
        {
            string propertiesPath = ConverterHelper.EmbeddedEngineProperties;
            if (!File.Exists(propertiesPath))
            {
                throw new FileNotFoundException("Could not find embedded engine properties file at " + propertiesPath);
            }
            try
            {
                System.Xml.XmlDocument propsDoc = new System.Xml.XmlDocument();
                propsDoc.LoadXml(propertiesPath);
                XmlElement propertiesNode = propsDoc.GetElementsByTagName("properties")[0] as XmlElement;
                Items = EngineProperty.ListFromXml(propertiesNode);
            }
            catch (Exception ex)
            {
                throw new Exception("An error occured while loading embedded engine properties file", ex);
            }

        }
    }

    /// <summary>
    /// Utility class to manage the DAISY Pipeline engine, either embedded or app. <br/>
    /// </summary>
    public class Engine
    {
       
        public static void StopEmbeddedEngine(IConversionEventsHandler events = null)
        {
            var allProcesses = Process.GetProcessesByName("java")
                .Where((p) =>
                {
                    return p.MainModule.FileName.StartsWith(ConverterHelper.EmbeddedEnginePath);
                }).ToArray();
            if (allProcesses.Length > 0) {
                try {
                    allProcesses[0].Kill();
                } catch (Exception ex) {
                    AddinLogger.Error("Could not stop embedded engine",ex);
                }
            }
        }

    }
}
