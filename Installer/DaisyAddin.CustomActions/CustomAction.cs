using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;


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
			bool isSameArch = true;
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
				isSameArch = false;

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



			bool is64system = IntPtr.Size * 8 == 64 ? true : false;

			// FIXME Due to possible "Windows Apps" installation of word that does not provide an installation registry key, 
			// we use the office 2007 version number as default targeted version to install components for.
			session["LATESTWORDVERSION"] = lastWordVersionStr == String.Empty ? "13.0" : lastWordVersionStr;
			// Defaults to 32 bits on detection
			session["LATESTWORDISX64"] = is64system && (lastWordVersionStr != String.Empty) && isSameArch  ? "yes" : "no";


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

		private static readonly string DAISY_APP_URL = "https://github.com/daisy/pipeline-ui/releases/download/1.8.0/daisy-pipeline-setup-1.8.0.exe";

        /// <summary>
        /// Offer to download and install the daisy pipeline if it is not installed on the system.
        /// </summary>
        /// <param name="session"></param>
        /// <returns></returns>
        [CustomAction]
        public static ActionResult ProposeDPAppInstall(Session session)
        {
            session.Log("Check if daisy pipeline app is installed");
            RegistryKey softwareKeys = Registry.CurrentUser.OpenSubKey(@"Software");
			foreach(string subKey in softwareKeys.GetSubKeyNames()) {
                RegistryKey software = softwareKeys.OpenSubKey(subKey);
                if(software.GetValue("ShortcutName") != null && software.GetValue("ShortcutName").ToString() == "DAISY Pipeline") {
					session.Log("Daisy Pipeline app is already installed.");
					return ActionResult.Success; // Already installed, no need to propose installation
                }
            }

            session.Log("Daisy Pipeline app is not installed, proposing installation.");
            //IntPtr hdl = Process.GetCurrentProcess().MainWindowHandle;
            //// Crée une fenêtre WPF invisible
            //Window owner = new Window
            //{
            //    Width = 0,
            //    Height = 0,
            //    ShowInTaskbar = false,
            //    WindowStyle = WindowStyle.None,
            //    AllowsTransparency = true,
            //    Opacity = 0
            //};

            //// Attache la fenêtre au handle donné
            //var helper = new WindowInteropHelper(owner);
            //helper.Owner = hdl;

            //// Affiche la fenêtre invisible
            //owner.Show();

            //session.Log("Handle of the main window: " + hdl.ToString("X"));
            if (MessageBox.Show("SaveAsDAISY can now use the DAISY Pipeline app as backend for the conversions.\r\n" +
					"Do you want to download and install the DAISY Pipeline app now?\r\n", "Download DAISY Pipeline app", MessageBoxButton.YesNo,
					MessageBoxImage.Question
				) == MessageBoxResult.Yes
            ) {
				session.Log("User accepted to install the DAISY Pipeline app.");
                CustomActionProgress progress = new CustomActionProgress();
                progress.ProgressMessage.Content = "Downloading DAISY Pipeline app installer...";
				progress.Show();
				progress.Activate();
                progress.Focus();
                try {
                    // Download the install from DAISY Pipeline app
					string installerPath = Path.Combine(Path.GetTempPath(), "daisy-pipeline-setup.exe");
                    using (var client = new System.Net.WebClient()) {
                        session.Log("Downloading DAISY Pipeline app installer from " + DAISY_APP_URL);
                        ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
                        client.DownloadProgressChanged += (sender, e) => {
                            progress.DownloadProgressBar.Value = e.ProgressPercentage;
                            //progressBar.Value = e.ProgressPercentage;
                        };
                        Task t = client.DownloadFileTaskAsync(DAISY_APP_URL, installerPath);
                        while (!t.IsCompleted) {
                            Application.Current.Dispatcher.Invoke(
								System.Windows.Threading.DispatcherPriority.Background,
								new Action(delegate { })
							); // Keep the UI responsive
                        }
						progress.Close();
						progress = null;
                    }
    
                    var process = System.Diagnostics.Process.Start(installerPath);

				} catch (Exception ex) {
					session.Log("Error while launching DAISY Pipeline app installer: " + ex.Message);
					MessageBox.Show("An error occurred while trying to launch the DAISY Pipeline app installer.\r\n" +
						ex.Message + "\r\n" +
						"Please try to install it manually from \r\n" +
                        "https://github.com/daisy/pipeline-ui/releases/latest", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
				} finally {
					if(progress != null) progress.Close();
                }

            } else {
				session.Log("User declined to install the DAISY Pipeline app.");
            }
			//owner.Close();

            return ActionResult.Success;
        }

    }
}
