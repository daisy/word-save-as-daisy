using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using WixToolset.Dtf.WindowsInstaller;

namespace WixCustomActions
{
    public class CustomActions
    {
        [CustomAction]
        public static ActionResult DetectLatestMsWordVersion(Session session)
        {
            session.Log("Begin detect last MS Word version");
            try
            {
                // Search word on the current system architecture registry
                RegistryKey lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
                float lastOfficeVersion = 0.0f;
                float lastWordVersion = 0.0f;
                bool isSameArch = true;
                if (lKey != null)
                {
                    foreach (string subKey in lKey.GetSubKeyNames())
                    {
                        // Check if the key name is a version number
                        Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
                        Match result = versionNumber.Match(subKey);
                        if (result.Success)
                        {
                            session.Log(@"Found office " + result.Value + @" components in HKLM\SOFTWARE\Microsoft\Office");
                            // if it is a superior versionCheck if it has a word subkey
                            float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
                            if (lastOfficeVersion < version)
                            {
                                lastOfficeVersion = version;
                                RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
                                if (wordKey != null)
                                {
                                    lastWordVersion = version;
                                    session.Log("Found Word with office " + lastWordVersion);
                                }
                            }
                        }
                    }
                }
                else
                {
                    session.Log(@"no HKLM\SOFTWARE\Microsoft\Office key found in the registry.");
                }

                // Also search for an existing 32Bits office install for x64 system 
                // (it is often the default downloaded version for x64 system)
                lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office");
                float lastOffice32BitVersion = 0.0f;
                float lastWord32bitVersion = 0.0f;
                if (lKey != null)
                {
                    isSameArch = false;

                    foreach (string subKey in lKey.GetSubKeyNames())
                    {
                        // Check if the key name is a version number
                        Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
                        Match result = versionNumber.Match(subKey);
                        if (result.Success)
                        {
                            session.Log(@"Found office " + result.Value + @" components in HKLM\SOFTWARE\WOW6432Node\Microsoft\Office");
                            // if it is a superior versionCheck if it has a word subkey
                            float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
                            if (lastOffice32BitVersion < version)
                            {
                                lastOffice32BitVersion = version;
                                RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
                                if (wordKey != null)
                                {
                                    lastWord32bitVersion = version;
                                    session.Log("Found 32 bits version of Word with office " + result.Value);
                                }
                            }
                        }
                    }
                }
                else
                {
                    session.Log(@"no HKLM\SOFTWARE\WOW6432Node\Microsoft\Office key found in the registry.");
                }

                string lastWordVersionStr = lastWord32bitVersion > 0.0f || lastWordVersion > 0.0f ?
                    Math.Max(lastWord32bitVersion, lastWordVersion).ToString("F1", CultureInfo.InvariantCulture) :
                    String.Empty;



                bool is64system = IntPtr.Size * 8 == 64;

                // FIXME Due to possible "Windows Apps" installation of word that does not provide an installation registry key, 
                // we use the office 2007 version number as default targeted version to install components for.
                session["LATESTWORDVERSION"] = lastWordVersionStr == String.Empty ? "13.0" : lastWordVersionStr;
                // Defaults to 32 bits on detection
                session["LATESTWORDISX64"] = is64system && (lastWordVersionStr != String.Empty) && isSameArch ? "yes" : "no";


                if (lastWordVersionStr != string.Empty)
                    session.Log($"MS Word version ({lastWordVersionStr}) was detected");
                else
                    session.Log("Can not detect MS Word, using version 13.0 / 2007 as default target.");

                return ActionResult.Success;
            }
            catch (Exception ex)
            {
                session.Log(ex.ToString());
                return ActionResult.Failure;
            }
            
        }

        private static readonly Regex versionNumber = new Regex("(?<maj>[0-9]+)\\.(?<min>[0-9]+)\\.(?<patch>[0-9]+).*");

        private static readonly int minVersion = 10000 * 1 + 100 * 11 + 0; // version 1.11.0

        [CustomAction]
        public static ActionResult DownloadAndLaunchDPApp(Session session)
        {
            if (session.Features["DownloadAndLaunchDPApp"].RequestState == InstallState.Local)
            {
                // Check SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
                //RegistryKey lKeyApp = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                using (RegistryKey lKeyApp = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
                {
                    foreach (string subKey in lKeyApp.GetSubKeyNames())
                    {
                        RegistryKey appKey = lKeyApp.OpenSubKey(subKey);
                        string displayName = appKey.GetValue("DisplayName")?.ToString() ?? "";
                        string displayVersion = appKey.GetValue("DisplayVersion")?.ToString() ?? "";

                        if (displayName.StartsWith("DAISY Pipeline") && versionNumber.Match(displayVersion) is Match m && m.Success)
                        {
                            // computing a simple version number to compare versions (major*10000 + minor*100 + patch)
                            // (assumes that major, minor and patch are all in range [0..99])
                            int versionNumberValue = int.Parse(m.Groups["maj"].Value) * 10000 +
                                                int.Parse(m.Groups["min"].Value) * 100 +
                                                int.Parse(m.Groups["patch"].Value);
                            // Avoid installing the app if a compatible version is already installed (version 1.11.0 or superior)
                            if (versionNumberValue >= minVersion)
                            {
                                using (Record rec = new Record(3))
                                {
                                    rec.SetString(1, "DownloadAndLaunchDPApp");
                                    rec.SetString(2, "Compatible DAISY Pipeline App already installed");
                                    session.Message(InstallMessage.ActionStart, rec);
                                    return ActionResult.Success;
                                }
                            }
                        }
                    }
                }
                
                try
                {
                    using (RegistryKey lKeyApp = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"))
                    {
                        foreach (string subKey in lKeyApp.GetSubKeyNames())
                        {
                            RegistryKey appKey = lKeyApp.OpenSubKey(subKey);
                            string displayName = appKey.GetValue("DisplayName")?.ToString() ?? "";
                            string displayVersion = appKey.GetValue("DisplayVersion")?.ToString() ?? "";

                            if (displayName.StartsWith("DAISY Pipeline") && versionNumber.Match(displayVersion) is Match m && m.Success)
                            {
                               // computing a simple version number to compare versions (major*10000 + minor*100 + patch)
                                // (assumes that major, minor and patch are all in range [0..99])
                                int versionNumberValue = int.Parse(m.Groups["maj"].Value) * 10000 +
                                                    int.Parse(m.Groups["min"].Value) * 100 +
                                                    int.Parse(m.Groups["patch"].Value);
                                // Avoid installing the app if a compatible version is already installed (version 1.11.0 or superior)
                                if (versionNumberValue >= minVersion)
                                {
                                    using (Record rec = new Record(3))
                                    {
                                        rec.SetString(1, "DownloadAndLaunchDPApp");
                                        rec.SetString(2, "Compatible DAISY Pipeline App already installed");
                                        session.Message(InstallMessage.ActionStart, rec);
                                        return ActionResult.Success;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    session.Log("Could not look into the HKLM registry, proceed with user-based dpapp installation... (ex : " + ex.ToString() + ")");
                }
                
                using (Record rec = new Record(3))
                {
                    rec.SetString(1, "DownloadAndLaunchDPApp");
                    rec.SetString(2, "Downloading the DAISY Pipeline app installer...");
                    session.Message(InstallMessage.ActionStart, rec);
                }

                using (Record rec = new Record(3))
                {
                    rec.SetInteger(1, 0);
                    rec.SetInteger(2, 100);
                    session.Message(InstallMessage.Progress, rec);
                }

                try
                {
                    // Download the install from DAISY Pipeline app
                    string installerPath = Path.Combine(Path.GetTempPath(), "daisy-pipeline-setup.exe");
                    if (File.Exists(installerPath)) File.Delete(installerPath);
                    using (var client = new System.Net.WebClient())
                    {
                        ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
                        client.DownloadFile(
                            "https://github.com/daisy/pipeline-ui/releases/download/1.11.0/daisy-pipeline-setup-1.11.0-mistral.exe",
                            installerPath);
                        using (Record rec = new Record(3))
                        {
                            rec.SetInteger(1, 33);
                            rec.SetInteger(2, 100);
                            session.Message(InstallMessage.Progress, rec);
                        }
                    }
                    using (Record rec = new Record(3))
                    {
                        rec.SetString(1, "DownloadAndLaunchDPApp");
                        rec.SetString(2, "Silent installation of DAISY Pipeline app...");
                        session.Message(InstallMessage.ActionStart, rec);
                    }
                    using (Record rec = new Record(3))
                    {
                        rec.SetInteger(1, 50);
                        rec.SetInteger(2, 100);
                        session.Message(InstallMessage.Progress, rec);
                    }
                    ProcessStartInfo exec = new ProcessStartInfo()
                    {
                        FileName = installerPath,
                        UseShellExecute = true,
                        Arguments = "/S" // silent install
                    };
                    var process = System.Diagnostics.Process.Start(exec);
                    using (Record rec = new Record(3))
                    {
                        rec.SetInteger(1, 75);
                        rec.SetInteger(2, 100);
                        session.Message(InstallMessage.Progress, rec);
                    }
                    process.WaitForExit();
                    using (Record rec = new Record(3))
                    {
                        rec.SetString(1, "DownloadAndLaunchDPApp");
                        rec.SetString(2, "Silent installation of DAISY Pipeline app completed.");
                        session.Message(InstallMessage.ActionStart, rec);
                    }
                    using (Record rec = new Record(3))
                    {
                        rec.SetInteger(1, 100);
                        rec.SetInteger(2, 100);
                        session.Message(InstallMessage.Progress, rec);
                    }
                    return ActionResult.Success;
                }
                catch (Exception ex)
                {
                    session.Log("Could not install DPApp : " + ex.ToString());
                    //return ActionResult.Failure;
                    //MessageBox.Show("An error occurred while trying to launch the DAISY Pipeline app installer.\r\n" +
                    //    ex.Message + "\r\n" +
                    //    "Please try to install it manually from \r\n" +
                    //    "https://github.com/daisy/pipeline-ui/releases/latest", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                session.Log("DownloadAndLaunchDPApp feature is not selected, skipping.");
                
            }
            return ActionResult.Success;

        }
    }
}
