using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

// PLEASE DEFINE THE FOLLOWING COMPILATION SYMBOL 
// X64INSTALLER : create the installer for Office 64 bits only
// UNIFIED : create the installer for both version on office

namespace DaisyInstaller
{
    // possibility from DAISY Meeting:
    // The installer could download the msi package from a web location instead of embedding it

    static class Program
    {
        /* Versions of offices to use for minimal and maximal version support
         * Office 3.0 (Word 2.0c, Excel 4.0a, PowerPoint 3.0, Mail)
         * Office 4.0 (Word 6.0, Excel 4.0, PowerPoint 3.0)
         * Office 4.2 (Word 6.0, Excel 5.0, PowerPoint 4.0, « Microsoft Office Manager »)
         * Office 4.3 (Word 6.0, Excel 5.0, PowerPoint 4.0, Pro:Access 2.0)
         * Office 95/7.0 (Word 95, etc.)
         * Office 97/8.0 (Word 97, etc.)
         * Office 2000/9.0 (Word 2000, etc.)
         * Office XP/10.0 (Word 2002, etc.)
         * Office 2003/11.0 (Word 2003, etc.)
         * Office 2007/12.0 (Word 2007, etc.)
         * Office 2010/14.0 (Word 2010, etc.)
         * Office 2013/15.0 (Word 2013, etc.)
         * Office 2016/16.0 (Word 2016, etc.)
         * Office 2019/17.0 (Word 2019, etc.)
         */

        public static float minimalVersionSupport = 11.0f;
        public static float maximalVersionSupport = 17.0f;

        private static readonly string appVersion = "1.10.0";
        private static readonly string DAISY_APP_URL = $"https://github.com/daisy/pipeline-ui/releases/download/{appVersion}/daisy-pipeline-setup-{appVersion}.exe";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            
#if X64INSTALLER // only
            bool installerIsForOffice32Bits = false;
#else // x86 only installer or unified installer default version installed
            bool installerIsForOffice32Bits = true;
#endif

            // If we want to check for Windows Arch, but we assume that windows is x64 as Microsoft is pushing to remove Windows x86 release.
            int archPtrBitSize = IntPtr.Size * 8; // 32 or 64, depending on executing arch

            bool officeIs64bits = (archPtrBitSize == 64); // We assume by default that office x64 is probable on x64 system

            RegistryKey lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office");
            RegistryKey wordRoot = null;
            float lastVersion = 0.0f;
            foreach (string subKey in lKey.GetSubKeyNames()) {
                // Check if the key name is a version number
                Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
                Match result = versionNumber.Match(subKey);
                if (result.Success) {
                    // if it is a superior versionCheck if it has a word subkey
                    float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
                    if (lastVersion < version) {
                        lastVersion = version;
                        RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
                        if (wordKey != null) wordRoot = wordKey;
                    }
                }
            }
            // Check for 32bits install on x64 system
            if(archPtrBitSize == 64 && wordRoot == null) {
                officeIs64bits = false;
                lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office");
                lastVersion = 0.0f;
                foreach (string subKey in lKey.GetSubKeyNames()) {
                    // Check if the key name is a version number
                    Regex versionNumber = new Regex("[0-9]+\\.[0-9]+");
                    Match result = versionNumber.Match(subKey);
                    if (result.Success) {
                        // if it is a superior versionCheck if it has a word subkey
                        float version = float.Parse(result.Value, CultureInfo.InvariantCulture.NumberFormat);
                        if (lastVersion < version) {
                            lastVersion = version;
                            RegistryKey wordKey = lKey.OpenSubKey(subKey + @"\Word\InstallRoot");
                            if (wordKey != null) wordRoot = wordKey;
                        }
                    }
                }
            }
            string warning = "";
            string error = "";
            bool keepInstall = true;
            if (wordRoot == null) {
                InstallArchSelector userArchSelector = new InstallArchSelector();
                if(userArchSelector.ShowDialog() == DialogResult.OK) {
                    officeIs64bits = (userArchSelector.SelectedArchitecture == 64);
                } else {
                    return; // abort installation
                }
                //warning = "Microsoft Word was not found in your system registry.\r\nDo you want to continue anyway and install the addin for Office " + (installerIsForOffice32Bits ? "32Bits" : "64Bits") + "?" ;
                //warning += "\r\n\r\nPlease check your office \"bit\" version to ensure you have the correct installer (link is in your clipboard):\r\n https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19";
                //Clipboard.SetText("https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19");
            } 
            //else if (!(installerIsForOffice32Bits ^ officeIs64bits)) {
            //    error = "This installer is for Office " + (installerIsForOffice32Bits ? "32Bits" : "64Bits") + " while Office " + (officeIs64bits ? "64Bits" : "32Bits") + " was found on your system.\r\nPlease download the installer for Office " + (officeIs64bits ? "64Bits" : "32Bits") + ".";
            //} 
            else if (lastVersion < minimalVersionSupport || lastVersion > maximalVersionSupport) {
                warning = "This addin officially supports Microsoft Word from Office 2003, up to Office 2016.\r\nA newer version of word has been found on your system but may not load this addin correctly.\r\nDo you want to continue anyway ?";
            }

            if(error.Length > 0) {
                MessageBox.Show(error, "Wrong installer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                keepInstall = false;
            } else if (warning.Length > 0) {
                keepInstall = MessageBox.Show(warning, "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes;
            }

            // Launch the install process
            if (keepInstall) {
                /* based on my search : 
                 * it is not possible to use a single MSI to choose if the install should be in x86 or x64
                 * an MSI is tagged with a target architecture preventing a x64 msi to write to the x86 program files and registry
                 * For now i bundle both "raw" MSI packages of previous installers but
                 * if we want to reduce the installer size, i need to split the package in 3 : 
                 * - the x86 msi package with x86 jre to deploy in pipeline-lite
                 * - the x64 msi package with x64 jre to deploy in pipeline-lite
                 * - the pipeline cab file (without jre embedded but with JRE 64 or 32 bits search) referenced by both MSI as external cab
                 */

                try {

                    string tempPath = Path.GetTempPath();

                    // unpackage the msi needed for install
                    string daisySetupPath = Path.Combine(tempPath, "DaisyAddinForWordSetup.msi");
                    string pipelineCabPath = Path.Combine(tempPath, "pipeline.cab");
                    if (File.Exists(daisySetupPath)) File.Delete(daisySetupPath);
                    if (File.Exists(pipelineCabPath)) File.Delete(pipelineCabPath);
#if UNIFIED
                    File.WriteAllBytes(daisySetupPath, officeIs64bits ? 
                        DaisyInstaller.Properties.Resources.DaisyAddinForWordSetup_x64 :
                        DaisyInstaller.Properties.Resources.DaisyAddinForWordSetup_x86);
                    File.WriteAllBytes(pipelineCabPath, DaisyInstaller.Properties.Resources.pipeline_cab);
#elif X64INSTALLER
                    File.WriteAllBytes(daisySetupPath, DaisyInstaller.Properties.Resources.DaisyAddinForWordSetup_x64);
                    File.WriteAllBytes(pipelineCabPath, DaisyInstaller.Properties.Resources.pipeline_cab);
#else
                    File.WriteAllBytes(daisySetupPath, DaisyInstaller.Properties.Resources.DaisyAddinForWordSetup_x86);
                    File.WriteAllBytes(pipelineCabPath, DaisyInstaller.Properties.Resources.pipeline_cab);
#endif

                    // launch the msi
                    var process = Process.Start(daisySetupPath);
                    process.WaitForExit();
                    bool installApp = true;
                    // Check SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
                    RegistryKey lKeyApp = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                    foreach (string subKey in lKeyApp.GetSubKeyNames()) {
                        RegistryKey appKey = lKeyApp.OpenSubKey(subKey);
                        string displayName = appKey.GetValue("DisplayName")?.ToString() ?? "";
                        string displayVersion = appKey.GetValue("DisplayVersion")?.ToString() ?? "";
                        if (
                            displayName.StartsWith("DAISY Pipeline") && displayVersion == appVersion
                        ) {
                            installApp = false;
                            break;
                        }
                    }
                    if (installApp) {
                        if (MessageBox.Show($"SaveAsDAISY can use the DAISY Pipeline app {appVersion} as backend for the conversions.\r\n" +
                                "Do you want to download and install the DAISY Pipeline app now?\r\n", "Download DAISY Pipeline app", MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question
                            ) == DialogResult.Yes
                        ) {
                            Form progressDialog = new Form();
                            progressDialog.Text = "Download DAISY Pipeline app";
                            progressDialog.Width = 300;
                            progressDialog.Height = 100;
                            FlowLayoutPanel container = new FlowLayoutPanel() { 
                                Dock = DockStyle.Fill, 
                                AutoSize = true, 
                                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                                AutoScroll = true,
                                FlowDirection = FlowDirection.TopDown,
                                WrapContents = false,
                            };
                            progressDialog.Controls.Add(container);
                            
                            Label label = new Label() { Text = "Downloading DAISY Pipeline app installer, please wait...", AutoSize = true, Dock = DockStyle.Top };
                            //ProgressBar progressBar = new ProgressBar() { Dock = DockStyle.Bottom, Width = 290 };
                            //progressBar.Maximum = 100;

                            container.Controls.Add(label);
                            //container.Controls.Add(progressBar);

                            //CustomActionProgress progress = new CustomActionProgress();
                            progressDialog.Show();
                            progressDialog.Refresh();
                            progressDialog.Activate();
                            progressDialog.Focus();
                            try {
                                // Download the install from DAISY Pipeline app
                                string installerPath = Path.Combine(Path.GetTempPath(), "daisy-pipeline-setup.exe");
                                if(File.Exists(installerPath)) File.Delete(installerPath);
                                using (var client = new System.Net.WebClient()) {
                                    ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
                                    client.DownloadFile(DAISY_APP_URL, installerPath);
                                    
                                }
                                label.Text = "Installing DAISY Pipeline app...";
                                ProcessStartInfo exec = new ProcessStartInfo() {
                                    FileName = installerPath,
                                    UseShellExecute = true,
                                    Arguments = "/S" // silent install
                                };
                                process = System.Diagnostics.Process.Start(exec);
                                process.WaitForExit();
                                label.Text = "DAISY Pipeline app is installed";
                                Thread.Sleep(3000);
                                progressDialog.Close();
                                progressDialog = null;

                            }
                            catch (Exception ex) {
                                MessageBox.Show("An error occurred while trying to launch the DAISY Pipeline app installer.\r\n" +
                                    ex.Message + "\r\n" +
                                    "Please try to install it manually from \r\n" +
                                    "https://github.com/daisy/pipeline-ui/releases/latest", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            } finally {
                                if (progressDialog != null) progressDialog.Close();
                            }

                        } else {
                            //session.Log("User declined to install the DAISY Pipeline app.");
                        }
                    }

                }
                catch (Exception e) {
                    MessageBox.Show(
                        "An error occured while installing the SaveAsDAISY add-in." +
                            "\r\n" + e.Message + "\r\n" + e.StackTrace,
                        "Install error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error
                        );
                    throw;
                }



            } else return;
            
        }
    }
    
}



