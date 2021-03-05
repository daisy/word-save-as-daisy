using System.Collections;
using Daisy.DaisyConverter.DaisyConverterLib;

namespace DaisyWord2007AddIn
{
	public class PreparetionResult
	{
		public PreparetionResult()
		{
			ObjectShapes = new ArrayList();
			ImageId = new ArrayList();
			InlineShapes = new ArrayList();
			InlineId = new ArrayList();
			//MultipleMathMl = new Hashtable();
			//MultipleOwnMathMl = new Hashtable();
		}

		public static PreparetionResult Failed(string error)
		{
			return new PreparetionResult() { IsSuccess = false, IsCanceled = false, LastMessage = error };
		}

		public static PreparetionResult Canceled(string message) {
			return new PreparetionResult() { IsSuccess = false, IsCanceled = true, LastMessage = message };
		}

		public bool IsSuccess { get; set; }

		public bool IsCanceled { get; set; }

		public string LastMessage { get; set; }
		public string OriginalFilePath { get; set; }
		public string TempFilePath { get; set; }
		public Initialize InitializeWindow { get; set; }
		public string MasterSubFlag { get; set; }
		public ArrayList ObjectShapes { get; set; }
		public ArrayList ImageId { get; set; }
		public ArrayList InlineShapes { get; set; }
		public ArrayList InlineId { get; set; }
		//public Hashtable MultipleMathMl { get; set; }
		//public Hashtable MultipleOwnMathMl { get; set; }
	}
}