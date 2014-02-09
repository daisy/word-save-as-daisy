using System.IO;
using System.Xml;

namespace Sonata.DaisyConverter.DaisyConverterLib
{
	/// <summary>
	/// Provides access to prepopulated_daisy xml file from AppData\..\Sonata directory.
	/// </summary>
	public class PrepopulateDaisyXml
	{
		/// <summary>
		/// Path for prepopulated_daisy xml file.
		/// </summary>
		private string PathForXml
		{
			get { return AddInHelper.AppDataSonataDirectory + "\\prepopulated_daisy.xml"; }
		}

		/// <summary>
		/// Creates new instance <see cref="PrepopulateDaisyXml"/> and 
		/// read publisher/title/creator if xml file exists.
		/// </summary>
		public PrepopulateDaisyXml()
		{
			Publisher = string.Empty;
			Title = string.Empty;
			Publisher = string.Empty;

			if (File.Exists(PathForXml))
			{
				XmlDocument document = new XmlDocument();
				document.Load(PathForXml);
				Creator = document.FirstChild.ChildNodes[0].InnerText;
				Title = document.FirstChild.ChildNodes[1].InnerText;
				Publisher = document.FirstChild.ChildNodes[2].InnerText;
			}
		}

		/// <summary>
		/// Save current publisher/title/creator values to xml file.
		/// </summary>
		public void Save()
		{
			XmlDocument docuemnt = new XmlDocument();

			XmlElement elmtDaisy = docuemnt.CreateElement("Daisy");
			docuemnt.AppendChild(elmtDaisy);

			XmlElement elmtCreator, elmtTitle, elmtPublisher;

			elmtCreator = docuemnt.CreateElement("Creator");
			elmtDaisy.AppendChild(elmtCreator);
			elmtCreator.InnerText = Creator;

			elmtTitle = docuemnt.CreateElement("Title");
			elmtDaisy.AppendChild(elmtTitle);
			elmtTitle.InnerText = Title;

			elmtPublisher = docuemnt.CreateElement("Publisher");
			elmtDaisy.AppendChild(elmtPublisher);
			elmtPublisher.InnerText = Publisher;

			if (!Directory.Exists(AddInHelper.AppDataSonataDirectory))
				Directory.CreateDirectory(AddInHelper.AppDataSonataDirectory);
			
			docuemnt.Save(PathForXml);
		}

		/// <summary>
		/// Publisher.
		/// </summary>
		public string Publisher { get; set; }

		/// <summary>
		/// Title.
		/// </summary>
		public string Title { get; set; }

		/// <summary>
		/// Creator.
		/// </summary>
		public string Creator { get; set; }
	}
}