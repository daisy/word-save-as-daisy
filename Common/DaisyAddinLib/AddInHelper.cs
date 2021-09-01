using Daisy.SaveAsDAISY.Conversion;
using System;
using System.IO;
using System.Reflection;

namespace Daisy.SaveAsDAISY.Forms
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
		public static bool buttonIsSingleWordToXMLConversion(string buttonName)
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
			return buttonIsSingleWordToXMLConversion(buttonName) || !ConverterHelper.PipelineIsInstalled();
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


		public static System.Resources.ResourceManager LabelsManager = new System.Resources.ResourceManager(
			"DaisyAddinLib.resources.Labels",
			Assembly.GetExecutingAssembly()
		);
        
		
		
	}
}