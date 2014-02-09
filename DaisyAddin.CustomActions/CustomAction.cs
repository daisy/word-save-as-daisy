using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;


namespace DaisyAddin.CustomActions
{
	public class CustomActions
	{
		[CustomAction]
		public static ActionResult DetectLatestMsWordVersion(Session session)
		{
			session.Log("Begin detect last MS Word version");

            string latestMSWordVersion = GetLatestMsWordVersion();
            session["LATESTWORDVERSION"] = latestMSWordVersion;

            if (latestMSWordVersion != string.Empty)
				session.Log("Latest MS Word version was detected");
			else
				session.Log("Can not detect MS Word");

			return ActionResult.Success;
		}
        
		private static string GetLatestMsWordVersion()
		{
			if (IsRegistryKeyExists(@"SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot"))
				return "14.0";

            if (IsRegistryKeyExists(@"SOFTWARE\Microsoft\Office\12.0\Word\InstallRoot"))
				return "12.0";

            if (IsRegistryKeyExists(@"SOFTWARE\Microsoft\Office\11.0\Word\InstallRoot"))
				return "11.0";

            if (IsRegistryKeyExists(@"SOFTWARE\Microsoft\Office\10.0\Word\InstallRoot"))
				return "10.0";

			return string.Empty;
		}

		private static bool IsRegistryKeyExists(string keyPath)
		{
			RegistryKey lKey = Registry.LocalMachine.OpenSubKey(keyPath);

			return lKey != null;
		}
	}
}
