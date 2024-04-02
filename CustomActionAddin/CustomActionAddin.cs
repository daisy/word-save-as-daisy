using System;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;

namespace CustomActionAddin
{
    [RunInstaller(true)]
    public class CustomActionAddin : Installer
    {
        /// <summary>
        /// Overrides Installer.Uninstall, which will be executed during uninstall process.
        /// </summary>
        /// <param name="savedState">The saved state.</param>
        public override void Uninstall(IDictionary savedState)
        {
            try
            {
                string fileName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                DeleteAllFiles(fileName);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Function To remove all the files from the Application Folder.
        /// </summary>
        /// <param name="path"></param>
        public void DeleteAllFiles(string path)
        {
            string[] fileEntries = Directory.GetFiles(path);
            string CustomActionPath = path + "\\CustomActionAddin.dll";
            foreach (string filename in fileEntries)
            {
                if (filename != CustomActionPath)
                {
                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }
                }
            }
            // Removing pipeline lite ms if it exists
            if (Directory.Exists(path + "\\pipeline-lite-ms"))
                Directory.Delete(path + "\\pipeline-lite-ms", true);
            // Remove pipeline 2
            if (Directory.Exists(path + "\\daisy-pipeline"))
                Directory.Delete(path + "\\daisy-pipeline", true);

        }
    }
}
