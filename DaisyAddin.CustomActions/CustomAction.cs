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

			// Search word on the current system architecture registry
			RegistryKey lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
			float lastOfficeVersion = 0.0f;
			float lastWordVersion = 0.0f;
			if (lKey != null) {
				foreach (string subKey in lKey.GetSubKeyNames()) {
					// Check if the key name is a version number
					Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
					Match result = versionNumber.Match(subKey);
					if (result.Success) {
						session.Log(@"Found office " + result.Value + @" components in HKLM\SOFTWARE\Microsoft\Office");
						// if it is a superior versionCheck if it has a word subkey
						float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
						if (lastOfficeVersion < version) {
							lastOfficeVersion = version;
							RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
							if (wordKey != null) {
								lastWordVersion = version;
								session.Log("Found Word with office " + lastWordVersion);
							}
						}
					}
				}
			} else {
				session.Log(@"no HKLM\SOFTWARE\Microsoft\Office key found in the registry.");
			}
			
			// Also search for an existing 32Bits office install for x64 system 
			// (it is often the default downloaded version for x64 system)
			lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office");
			float lastOffice32BitVersion = 0.0f;
			float lastWord32bitVersion = 0.0f;
			if (lKey != null) {
				foreach (string subKey in lKey.GetSubKeyNames()) {
					// Check if the key name is a version number
					Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
					Match result = versionNumber.Match(subKey);
					if (result.Success) {
						session.Log(@"Found office " + result.Value + @" components in HKLM\SOFTWARE\WOW6432Node\Microsoft\Office");
						// if it is a superior versionCheck if it has a word subkey
						float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
						if (lastOffice32BitVersion < version) {
							lastOffice32BitVersion = version;
							RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
							if (wordKey != null) {
								lastWord32bitVersion = version;
								session.Log("Found 32 bits version of Word with office " + result.Value);
							}
						}
					}
				}
			} else {
				session.Log(@"no HKLM\SOFTWARE\WOW6432Node\Microsoft\Office key found in the registry.");
			}

			string lastWordVersionStr = lastWord32bitVersion > 0.0f || lastWordVersion > 0.0f ?
				Math.Max(lastWord32bitVersion, lastWordVersion).ToString("F1", CultureInfo.InvariantCulture) :
				String.Empty;




			// FIXME Due to possible "Windows Apps" installation of word that does not provide an installation registry key, 
			// we use the office 2007 version number as default targeted version to install components for.
			session["LATESTWORDVERSION"] = lastWordVersionStr == String.Empty ? "13.0" : lastWordVersionStr;
			

            if (lastWordVersionStr != string.Empty)
				session.Log($"MS Word version ({lastWordVersionStr}) was detected");
			else
				session.Log("Can not detect MS Word, using version 13.0 / 2007 as default target.");

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
