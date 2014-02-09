using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;

namespace DaisyInstaller
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // search of MS Word intalling last version
            RegistryKey lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot");
            if (lKey == null)
                lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\Word\InstallRoot");
            if (lKey == null)
                lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Word\InstallRoot");
            if (lKey == null)
                lKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\10.0\Word\InstallRoot");

            if (lKey == null)
            {
                MessageBox.Show("Microsoft Word do not install.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // start DaisyAddinForWordSetup.exe
            string daisySetupPath = Path.Combine(Path.GetTempPath(), "DaisyAddinForWordSetup.exe");
            File.WriteAllBytes(daisySetupPath, Properties.Resources.DaisyAddinForWordSetup);

            Process.Start(daisySetupPath);
        }
    }
}
