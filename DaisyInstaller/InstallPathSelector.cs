using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DaisyInstaller {
    public partial class InstallPathSelector : Form {

        
        
        private string defaultRoot; // program files 
        private string[] defaultTree; // folder tree under defaultRoot
        private string[] installDefaultFullPath; // defaultRoot + defaultTree

        public InstallPathSelector(string defaultRoot = null, string[] defaultTree = null) {
            InitializeComponent();
            this.defaultRoot = defaultRoot;
            this.defaultTree = defaultTree;

            this.installDefaultFullPath = new string[defaultTree.Length + 1];
            installDefaultFullPath[0] = defaultRoot;
            defaultTree.CopyTo(installDefaultFullPath, 1);

            this.SelectedInstallDir.Text = System.IO.Path.Combine(installDefaultFullPath);
            Directory.CreateDirectory(this.SelectedInstallDir.Text);
        }

        private void ChangeInstallDir_Click(object sender, EventArgs e) {
            FolderBrowserDialog installDirSelection = new FolderBrowserDialog();
            installDirSelection.SelectedPath = this.SelectedInstallDir.Text;

            if(installDirSelection.ShowDialog() == DialogResult.OK) {
                this.SelectedInstallDir.Text = installDirSelection.SelectedPath;
            }

        }

        public string getInstallDir() {
            return this.SelectedInstallDir.Text;
        }

        private void Accept_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.OK;
            // Cleanup install tree if directory has changed
            if (this.SelectedInstallDir.Text != System.IO.Path.Combine(installDefaultFullPath)) {
                string currentDeletablePath = System.IO.Path.Combine(installDefaultFullPath);
                for (int index = this.defaultTree.Length - 1; index >= 0; --index) {
                    try {
                        Directory.Delete(currentDeletablePath);
                    } catch (Exception ex) {
                        Console.WriteLine(ex.Message);
                    }
                    currentDeletablePath = Directory.GetParent(currentDeletablePath).FullName;
                }
            }
            this.Close();
        }

        private void Cancel_Click(object sender, EventArgs e) {
            this.DialogResult = DialogResult.Cancel;
            // Cleanup install tree
            string currentDeletablePath = System.IO.Path.Combine(installDefaultFullPath);
            for (int index = this.defaultTree.Length - 1; index >= 0; --index) {
                try {
                    Directory.Delete(currentDeletablePath);
                } catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
                currentDeletablePath = Directory.GetParent(currentDeletablePath).FullName;
            }
            this.Close();
        }

    }
}
