namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// Provides details about a conversion.
	/// </summary>
	public class ConversionResult
	{
		public enum Type {
			Success,
			Cancel,
			Error
		}

		public bool Canceled { get => ResultType == Type.Cancel; }
		public bool Succeeded { get => ResultType == Type.Success; }

        public bool Failed { get => ResultType == Type.Error; }

        private ConversionResult(Type resultType)
		{
			ResultType = resultType;
		}

		/// <summary>
		/// Creates success result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Success()
		{
			return new ConversionResult(Type.Success);
		}

		/// <summary>
		/// Creates cancel result.
		/// </summary>
		/// <returns></returns>
		public static ConversionResult Cancel()
		{
			return new ConversionResult(Type.Cancel);
		}


		/// <summary>
		/// Creates UnknownError result.
		/// </summary>
		/// <param name="errorMessage"></param>
		/// <returns></returns>
		public static ConversionResult Fail(string errorMessage)
		{
            return new ConversionResult(Type.Error) {
				ErrorMessage = errorMessage
			};
		}

		/// <summary>
		/// Error message.
		/// </summary>
		public string ErrorMessage { get; set; }


		/// <summary>
		/// Result type.
		/// </summary>
		public Type ResultType { get; set; }
	}
}