namespace Daisy.SaveAsDAISY.Conversion.PageStylesValidation
{
	public class ValidatoinResult
	{
		public ValidatoinResult()
		{

		}

		public ValidatoinResult(string errorMessage)
		{
			ErrorMessage = errorMessage;
		}

		public bool IsValid { get { return string.IsNullOrEmpty(ErrorMessage); } }

		public string ErrorMessage { get; private set; }
	}
}