using System.IO;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// Provides access to prepopulated_daisy_output.xml from AppData\..\Sonanta directory.
	/// </summary>
	public class PrepopulateDaisyOutput
	{
		/// <summary>
		/// Path to xml file.
		/// </summary>
		private static string XmlFilePath
		{
			get { return ConverterHelper.AppDataSaveAsDAISYDirectory + "\\prepopulated_daisy_output.xml"; }
		}

		/// <summary>
		/// Create new instance of <see cref="PrepopulateDaisyOutput"/>.
		/// </summary>
		private PrepopulateDaisyOutput()
		{
			OutputPath = string.Empty;
		}

		/// <summary>
		/// Create new instance of <see cref="PrepopulateDaisyOutput"/> and set output path.
		/// </summary>
		/// <param name="outputPath">Output path.</param>
		public PrepopulateDaisyOutput(string outputPath)
		{
			OutputPath = outputPath;
		}

		/// <summary>
		/// Try to load data from prepolated_daisy_output.xml
		/// </summary>
		/// <returns>Instnace of <see cref="PrepopulateDaisyOutput"/> class if xml file exists, otherwise - null.</returns>
		public static PrepopulateDaisyOutput Load()
		{
			if( ! File.Exists(XmlFilePath))
				return null;

			PrepopulateDaisyOutput result = new PrepopulateDaisyOutput();;

			var document = new XmlDocument();
			document.Load(XmlFilePath);

			if (document.FirstChild.ChildNodes.Count > 0)
				result.OutputPath = document.FirstChild.ChildNodes[0].InnerText;

			return result;
		}

		/// <summary>
		/// Update existing or create new prepopulated_daisy_output.xml. OutputFile Path
		/// will be equal to <see cref="OutputPath"/>.
		/// </summary>
		public void Save()
		{
			var document = new XmlDocument();

			if (!Directory.Exists(ConverterHelper.AppDataSaveAsDAISYDirectory))
				Directory.CreateDirectory(ConverterHelper.AppDataSaveAsDAISYDirectory);

			if (!File.Exists(XmlFilePath))
			{
				var outputXmlDaisy = document.CreateElement("Daisy");
				document.AppendChild(outputXmlDaisy);

				var outputXmlFilePath = document.CreateElement("OutputFilePath");
				outputXmlFilePath.InnerText = OutputPath;
				outputXmlDaisy.AppendChild(outputXmlFilePath);
			}
			else
			{
				document.Load(XmlFilePath);
				document.FirstChild.ChildNodes[0].InnerText = OutputPath;
			}
            
			document.Save(XmlFilePath);
		}

		/// <summary>
		/// Output path.
		/// </summary>
		public string OutputPath { get; set; }

	}
}