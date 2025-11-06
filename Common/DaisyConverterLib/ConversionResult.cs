using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;

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
		public static ConversionResult Fail(Exception error)
		{
            return new ConversionResult(Type.Error) {
                ErrorDetails = error
            };
		}

		/// <summary>
		/// Error message.
		/// </summary>
		public Exception ErrorDetails { get; set; }

		public string ErrorMessage
		{
			get
			{
				if (ErrorDetails != null)
				{
					string message = ErrorDetails.Message;
					if(ErrorDetails is JobRequestError) {
						message += "\r\n" + ((JobRequestError)ErrorDetails).Description;
                    }
                    var inner = ErrorDetails.InnerException;
                    while (inner != null) {
                        message += "\r\n" + inner.Message;
						inner = inner.InnerException;
                    }
                    return message;
				}
				return null;
			}
        }


        /// <summary>
        /// Result type.
        /// </summary>
        public Type ResultType { get; set; }
	}
}