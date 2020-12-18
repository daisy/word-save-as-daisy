namespace Daisy.DaisyConverter.DaisyConverterLib
{
	/// <summary>
	/// Provides details about convertation.
	/// </summary>
	public class ConvertResult
	{
		private ConvertResult(ConvertResultType resultType)
		{
			ResultType = resultType;
		}

		/// <summary>
		/// Creates success result.
		/// </summary>
		/// <returns></returns>
		public static ConvertResult Success()
		{
			return new ConvertResult(ConvertResultType.Success);
		}

		/// <summary>
		/// Creates cancel result.
		/// </summary>
		/// <returns></returns>
		public static ConvertResult Cancel()
		{
			return new ConvertResult(ConvertResultType.Cancel);
		}

		/// <summary>
		/// Creates invalid result.
		/// </summary>
		/// <param name="error"></param>
		/// <returns></returns>
		public static ConvertResult ValidationError(string error)
		{
			var result = new ConvertResult(ConvertResultType.ValidationError);
			result.ValidationErrorMessage = error;
			return result;
		}

		/// <summary>
		/// Creates UnknownError result.
		/// </summary>
		/// <param name="errorMessage"></param>
		/// <returns></returns>
		public static ConvertResult UnknownError(string errorMessage)
		{
			var result = new ConvertResult(ConvertResultType.UnknownError);
			result.UnknownErrorMessage = errorMessage;
			return result;
		}

		/// <summary>
		/// Error message.
		/// </summary>
		public string UnknownErrorMessage { get; set; }

		/// <summary>
		/// Validation error message.
		/// </summary>
		public string ValidationErrorMessage { get; set; }

		/// <summary>
		/// Result type.
		/// </summary>
		public ConvertResultType ResultType { get; set; }
	}
}