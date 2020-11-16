namespace DaisyInstaller {
    partial class InstallPathSelector {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.SelectedInstallDir = new System.Windows.Forms.TextBox();
            this.ChangeInstallDir = new System.Windows.Forms.Button();
            this.Accept = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // SelectedInstallDir
            // 
            this.SelectedInstallDir.Location = new System.Drawing.Point(12, 36);
            this.SelectedInstallDir.Name = "SelectedInstallDir";
            this.SelectedInstallDir.Size = new System.Drawing.Size(368, 20);
            this.SelectedInstallDir.TabIndex = 0;
            // 
            // ChangeInstallDir
            // 
            this.ChangeInstallDir.Location = new System.Drawing.Point(386, 34);
            this.ChangeInstallDir.Name = "ChangeInstallDir";
            this.ChangeInstallDir.Size = new System.Drawing.Size(75, 23);
            this.ChangeInstallDir.TabIndex = 1;
            this.ChangeInstallDir.Text = "Change";
            this.ChangeInstallDir.UseVisualStyleBackColor = true;
            this.ChangeInstallDir.Click += new System.EventHandler(this.ChangeInstallDir_Click);
            // 
            // Accept
            // 
            this.Accept.Location = new System.Drawing.Point(386, 63);
            this.Accept.Name = "Accept";
            this.Accept.Size = new System.Drawing.Size(75, 23);
            this.Accept.TabIndex = 2;
            this.Accept.Text = "OK";
            this.Accept.UseVisualStyleBackColor = true;
            this.Accept.Click += new System.EventHandler(this.Accept_Click);
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(305, 63);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(75, 23);
            this.Cancel.TabIndex = 3;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(312, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Select the folder where the addin will be installed on your system.";
            // 
            // InstallPathSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 98);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.Accept);
            this.Controls.Add(this.ChangeInstallDir);
            this.Controls.Add(this.SelectedInstallDir);
            this.Name = "InstallPathSelector";
            this.Text = "Install folder selection";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox SelectedInstallDir;
        private System.Windows.Forms.Button ChangeInstallDir;
        private System.Windows.Forms.Button Accept;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.Label label1;
    }
}