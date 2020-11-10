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
        public static float maximalVersionSupport = 16.0f;

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
                warning = "This addin officially supports Microsoft Word from Office 2003, up to Office 2016.\r\nA newer version of word has beend found on your system but may not load this addin correctly.\r\nDo you want to continue anyway ?";
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
                 * - the pipeline zip file (without jre embedded but with JRE 64 or 32 bits search)
                 * The install directory need to be set by the user at the bootstraper stage (so in this project)
                 * and passed along to the "TARGETDIR" MSI property
                 * that could be done using the commande line 'msiexec /i x86_or_x84_msi TARGETDIR="C:\user_select_path"'
                 * Then we could finish by unzipping the pipeline in the selected directory
                 */


                // start DaisyAddinForWordSetup.msi
                string tempPath = Path.GetTempPath();
                string daisySetupPath = Path.Combine(tempPath, "DaisyAddinForWordSetup.msi");
#if UNIFIED
                if(officeIs64bits)
                    File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x64);
                else File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x86);
#elif X64INSTALLER
                File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x64);
#else
                File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup_x86);
#endif

                Process install = Process.Start(daisySetupPath);

            } else return;
            
        }
    }
    
    /* for future use with pipeline zip extracted from msi packages
     * (required as i am using .net 4 framework that i know is included by default with windows 10 
     * but does not directly "expose" the system API for zip handling in .net, and i don't want to use another third party lib)
    // ZipArchive from https://www.codeproject.com/Articles/209731/Csharp-use-Zip-archives-without-external-libraries
    class ZipArchive : IDisposable {
        private object external;
        private ZipArchive() { }
        public enum CompressionMethodEnum { Stored, Deflated };
        public enum DeflateOptionEnum { Normal, Maximum, Fast, SuperFast };
        //...
        public static ZipArchive OpenOnFile(string path, FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read, FileShare share = FileShare.Read, bool streaming = false) {
            var type = typeof(System.IO.Packaging.Package).Assembly.GetType("MS.Internal.IO.Zip.ZipArchive");
            var meth = type.GetMethod("OpenOnFile", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipArchive { external = meth.Invoke(null, new object[] { path, mode, access, share, streaming }) };
        }
        public static ZipArchive OpenOnStream(Stream stream, FileMode mode = FileMode.OpenOrCreate, FileAccess access = FileAccess.ReadWrite, bool streaming = false) {
            var type = typeof(System.IO.Packaging.Package).Assembly.GetType("MS.Internal.IO.Zip.ZipArchive");
            var meth = type.GetMethod("OpenOnStream", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipArchive { external = meth.Invoke(null, new object[] { stream, mode, access, streaming }) };
        }
        public ZipFileInfo AddFile(string path, CompressionMethodEnum compmeth = CompressionMethodEnum.Deflated, DeflateOptionEnum option = DeflateOptionEnum.Normal) {
            var type = external.GetType();
            var meth = type.GetMethod("AddFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            var comp = type.Assembly.GetType("MS.Internal.IO.Zip.CompressionMethodEnum").GetField(compmeth.ToString()).GetValue(null);
            var opti = type.Assembly.GetType("MS.Internal.IO.Zip.DeflateOptionEnum").GetField(option.ToString()).GetValue(null);
            return new ZipFileInfo { external = meth.Invoke(external, new object[] { path, comp, opti }) };
        }
        public void DeleteFile(string name) {
            var meth = external.GetType().GetMethod("DeleteFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            meth.Invoke(external, new object[] { name });
        }
        public void Dispose() {
            ((IDisposable)external).Dispose();
        }
        public ZipFileInfo GetFile(string name) {
            var meth = external.GetType().GetMethod("GetFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipFileInfo { external = meth.Invoke(external, new object[] { name }) };
        }

        public IEnumerable<ZipFileInfo> Files {
            get {
                var meth = external.GetType().GetMethod("GetFiles", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                var coll = meth.Invoke(external, null) as System.Collections.IEnumerable; //ZipFileInfoCollection
                foreach (var p in coll) yield return new ZipFileInfo { external = p };
            }
        }
        public IEnumerable<string> FileNames {
            get { return Files.Select(p => p.Name).OrderBy(p => p); }
        }

        public struct ZipFileInfo {
            internal object external;
            private object GetProperty(string name) {
                return external.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(external, null);
            }
            public override string ToString() {
                return Name;// base.ToString();
            }
            public string Name {
                get { return (string)GetProperty("Name"); }
            }
            public DateTime LastModFileDateTime {
                get { return (DateTime)GetProperty("LastModFileDateTime"); }
            }
            public bool FolderFlag {
                get { return (bool)GetProperty("FolderFlag"); }
            }
            public bool VolumeLabelFlag {
                get { return (bool)GetProperty("VolumeLabelFlag"); }
            }
            public object CompressionMethod {
                get { return GetProperty("CompressionMethod"); }
            }
            public object DeflateOption {
                get { return GetProperty("DeflateOption"); }
            }
            public Stream GetStream(FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read) {
                var meth = external.GetType().GetMethod("GetStream", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                return (Stream)meth.Invoke(external, new object[] { mode, access });
            }
        }
    }*/
}
