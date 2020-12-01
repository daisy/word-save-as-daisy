using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Security.Permissions;

namespace DaisyInstaller {

    
    public partial class PipelineInstaller : Form {
        private string zipPath;
        private string unzippingPath;
        
        public PipelineInstaller(string zipPath, string unzippingPath) {
            InitializeComponent();
            this.zipPath = zipPath;
            this.unzippingPath = unzippingPath;
        }

        
        private void PipelineInstaller_Shown(object sender, EventArgs e) {
            installProgress.Minimum = 0;
            installProgress.Value = 0;
            using (var archive = ZipArchive.OpenOnFile(zipPath)) {
                installProgress.Maximum = archive.Files.Count();
                int i = 0;
                foreach (var zippedFile in archive.Files) {
                    string path = System.IO.Path.Combine(unzippingPath, zippedFile.Name);
                    this.Unzipping.Text = "Unzipping " + path;
                    this.Unzipping.Refresh();
                    if (zippedFile.FolderFlag) {
                        Directory.CreateDirectory(path);
                    } else {
                        string dir = System.IO.Path.GetDirectoryName(path);
                        if (!Directory.Exists(dir)) {
                            Directory.CreateDirectory(dir);
                        }
                        FileStream unzipped = File.OpenWrite(path);
                        zippedFile.GetStream().CopyTo(unzipped);
                        unzipped.Close();
                        
                    }
                    installProgress.Value = ++i;
                }
            }
            this.Close();
        }


    }
}
