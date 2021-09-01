using System;
using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Layout;

namespace Daisy.SaveAsDAISY.Conversion
{
	public class AddinLogger
	{
		private static ILog _log;

		#region initialization

		static AddinLogger()
		{
			ConfigureLogger();
			_log = LogManager.GetLogger(typeof (AddinLogger));
		}

		private static void ConfigureLogger()
		{
			PatternLayout patternLayout = new PatternLayout
			                              	{
			                              		ConversionPattern =
			                              			"%date{yyyy-MM-dd HH:mm:ss}  %-5level - %message%newline%newline"
			                              	};
			RollingFileAppender appender = new RollingFileAppender
			                               	{
			                               		AppendToFile = true,
			                               		File = ConverterHelper.AppDataSaveAsDAISYDirectory + @"logs\addin.log",
			                               		ImmediateFlush = true,
			                               		Name = "AddinFileAppender",
			                               		RollingStyle = RollingFileAppender.RollingMode.Size,
			                               		MaxSizeRollBackups = 10,
			                               		MaximumFileSize = "10MB",
			                               		LockingModel = new FileAppender.MinimalLock(),
			                               		Layout = patternLayout
			                               	};

			patternLayout.ActivateOptions();
			appender.ActivateOptions();

			BasicConfigurator.Configure(appender);
		}

		#endregion

		public static void Error(string errorMessage)
		{
			_log.Error(errorMessage);
		}

		public static void Error(Exception ex)
		{
			_log.Error("Error", ex);
		}

		public static void Info(string messsage)
		{
			_log.Info(messsage);
		}
	}
}