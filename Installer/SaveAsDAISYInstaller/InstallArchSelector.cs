using System;
using System.Windows.Forms;

namespace DaisyInstaller
{
    public partial class InstallArchSelector : Form {

        public int SelectedArchitecture { get; private set; } = IntPtr.Size *8;
        public InstallArchSelector() {
            InitializeComponent();
        }

        private void InstallForOffice32_Click(object sender, EventArgs e) {
            SelectedArchitecture = 32;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void InstallForOffice64_Click(object sender, EventArgs e) {
            SelectedArchitecture = 64;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            System.Diagnostics.Process.Start("https://support.microsoft.com/en-us/office/about-office-what-version-of-office-am-i-using-932788b8-a3ce-44bf-bb09-e334518b8b19");
        }
    }
}
