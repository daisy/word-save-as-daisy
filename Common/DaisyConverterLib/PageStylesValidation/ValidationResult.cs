namespace Daisy.SaveAsDAISY.Conversion.PageStylesValidation
{
	public class ValidationResult
	{
		public ValidationResult()
		{

		}

		public ValidationResult(string errorMessage)
		{
			ErrorMessage = errorMessage;
		}

		public bool IsValid { get { return string.IsNullOrEmpty(ErrorMessage); } }

		public string ErrorMessage { get; private set; }
	}
}