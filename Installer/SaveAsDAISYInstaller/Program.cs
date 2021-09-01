using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Linq;

// PLEASE DEFINE THE FOLLOWING COMPILATION SYMBOL 
// X64INSTALLER : create the installer for Office 64 bits only
// UNIFIED : create the installer for both version on office

namespace DaisyInstaller
{
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
                warning = "Microsoft Word was not found in your system registry.\r\nDo you want to continue anyway and install the addin for Office " + (installerIsForOffice32Bits ? "32Bits" : "64Bits") + "?" ;
                warning += "\r\n\r\nPlease check your office \"bit\" version to ensure you have the correct installer (link is in your clipboard):\r\n https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19";
                Clipboard.SetText("https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19");
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
                        Properties.Resources.DaisyAddinForWordSetup_x64 :
                        Properties.Resources.DaisyAddinForWordSetup_x86);
                    File.WriteAllBytes(pipelineCabPath, Properties.Resources.pipeline_cab);
#elif X64INSTALLER
                    File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x64);
                    File.WriteAllBytes(pipelineCabPath, Properties.Resources.pipeline_cab);
#else
                    File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x86);
                    File.WriteAllBytes(pipelineCabPath, Properties.Resources.pipeline_cab);
#endif

                    // launch the msi
                    Process.Start(daisySetupPath);

                } catch (Exception e) {
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



