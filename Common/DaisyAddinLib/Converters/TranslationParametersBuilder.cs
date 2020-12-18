using System.Collections;
using Daisy.DaisyConverter.DaisyConverterLib;

namespace Daisy.DaisyConverter.DaisyConverterLib.Converters
{
	/// <summary>
	/// Translation parameters builder.
	/// </summary>
	public class TranslationParametersBuilder
	{
		public string OutputFile { get; set; }
		public string Title { get; set; }
		public string Creator { get; set; }
		public string Publisher { get; set; }
		public string UID { get; set; }
		public string TrackChangesFlage { get; set; }
		public string Version { get; set; }
		public string MasterSubFlag { get; set; }
		public string Subject { get; set; }

		#region fluent initialization

		public TranslationParametersBuilder WithOutputFile(string outputFile)
		{
			OutputFile = outputFile;
			return this;
		}

		public TranslationParametersBuilder WithTitle(string title)
		{
			Title = title;
			return this;
		}

		public TranslationParametersBuilder WithCreator(string creator)
		{
			Creator = creator;
			return this;
		}

		public TranslationParametersBuilder WithSubject(string subject)
		{
			Subject = subject;
			return this;
		}

		public TranslationParametersBuilder WithPublisher(string publisher)
		{
			Publisher = publisher;
			return this;
		}

		public TranslationParametersBuilder WithUID(string uid)
		{
			UID = uid;
			return this;
		}

		public TranslationParametersBuilder WithTrackChangesFlag(string trackChangesFlage)
		{
			TrackChangesFlage = trackChangesFlage;
			return this;
		}

		public TranslationParametersBuilder WithVersion(string version)
		{
			Version = version;
			return this;
		}

		public TranslationParametersBuilder WithMasterSubFlag(string masterSubFlag)
		{
			MasterSubFlag = masterSubFlag;
			return this;
		}

		#endregion

		public bool AddParameter(string name, string value, Hashtable parametersHolder)
		{
			if (value == null)
				return false;

			parametersHolder.Add(name, value);
			return true;
		}

		/// <summary>
		/// Build parameters.
		/// </summary>
		/// <returns>Hastable with keys equals to parameter name and values equals to value of appropriate parameter.</returns>
		public Hashtable BuildTranslationParameters()
		{
			Hashtable parameters = new Hashtable();
			AddParameter("outputfile", OutputFile, parameters);
			AddParameter("Title", Title, parameters);
			AddParameter("Creator", Creator, parameters);
			AddParameter("Publisher", Publisher, parameters);
			AddParameter("UID", UID, parameters);
			AddParameter("Subject", Subject, parameters);

			AddParameter("TRACK", TrackChangesFlage, parameters);
			AddParameter("Version", Version, parameters);

			AddParameter("MasterSub", MasterSubFlag, parameters);

			DaisyTranslationSettings daisySt = new DaisyTranslationSettings();
			string imgoption = daisySt.GetImageOption;
			string resampleValue = daisySt.GetResampleValue;
			string characterStyle = daisySt.GetCharacterStyle;
			string pagenumStyle = daisySt.GetPagenumStyle;
			if (imgoption != " ")
			{
				parameters.Add("ImageSizeOption", imgoption);
				parameters.Add("DPI", resampleValue);
			}
			if (characterStyle != " ")
			{
				parameters.Add("CharacterStyles", characterStyle);
			}
			if (pagenumStyle != " ")
			{
				parameters.Add("Custom", pagenumStyle);
			}

			return parameters;
		}
	}
}