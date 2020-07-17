using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
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
			RegistryKey lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
			float lastOfficeVersion = 0.0f;
			string lastWordVersion = String.Empty;
			foreach (string subKey in lKey.GetSubKeyNames()) {
				// Check if the key name is a version number
				Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
				Match result = versionNumber.Match(subKey);
				if (result.Success) {
					// if it is a superior versionCheck if it has a word subkey
					float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
					if (lastOfficeVersion < version) {
						lastOfficeVersion = version;
						RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
						if (wordKey != null) {
							lastWordVersion = result.Value;
						}
					}
				}
			}
			return lastWordVersion;
		}
		
		/*
		private static bool IsRegistryKeyExists(string keyPath)
		{
			RegistryKey lKey = Registry.LocalMachine.OpenSubKey(keyPath);

			return lKey != null;
		}*/

	}
}
