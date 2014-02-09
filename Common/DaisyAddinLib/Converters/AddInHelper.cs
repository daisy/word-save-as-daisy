using System;
using System.IO;
using System.Reflection;

namespace Sonata.DaisyConverter.DaisyConverterLib
{
	public class AddInHelper
	{
		#region constants

		public const int DaisyStylesCount = 38;

		#endregion
		
		/// <summary>
		/// Indicates if single daisy button was clicked or full daisy.
		/// </summary>
		/// <param name="buttonName"></param>
		/// <returns></returns>
		public static bool IsSingleDaisyButton(string buttonName)
		{
			return buttonName == "DaisySingle" || buttonName == "DaisyTabSingle" || 
			       buttonName == "Daisy" ||  buttonName == "Button1";
		}

		/// <summary>
		/// Indicates when should be OOX should be trasformed to daisy xml 
		/// without pipeline script exectution.
		/// </summary>
		/// <param name="buttonName">Clicked button name.</param>
		/// <returns></returns>
		public static bool IsSingleDaisyTranslate(string buttonName)
		{
			return IsSingleDaisyButton(buttonName) || !IsPipelineExists();
		}

		/// <summary>
		/// Indicates when single daisy from multiple docx button was clicked or full daisy.
		/// </summary>
		/// <param name="buttonName"></param>
		/// <returns></returns>
		public static bool IsSingleDaisyFromMultipleButton(string buttonName)
		{
			return buttonName == "DaisyMultiple" || buttonName == "DaisyTabMultiple" || buttonName == "Button2";
		}

		/// <summary>
		/// Gets path to pipeline root directory.
		/// </summary>
		public static string PathForPipeline
		{
			get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms"; }
		}

		/// <summary>
		/// Indicates if pipeline exists.
		/// </summary>
		/// <returns></returns>
		public static bool IsPipelineExists()
		{
			return Directory.Exists(PathForPipeline);
		}

		/// <summary>
		/// Gets path to Sanata directory in AppData.
		/// </summary>
		public static string AppDataSonataDirectory
		{
			get { return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Sonata"; }
		}

		
	}
}