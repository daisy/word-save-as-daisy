
namespace DaisyInstaller {
    partial class InstallArchSelector {
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
            this.label1 = new System.Windows.Forms.Label();
            this.InstallForOffice32 = new System.Windows.Forms.Button();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.InstallForOffice64 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoEllipsis = true;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(271, 52);
            this.label1.TabIndex = 1;
            this.label1.Text = "The \"bit-version\" of your MS Office version could not be\r\nidentified by the insta" +
    "ll process.\r\nPlease check your office bit-version and click on the\r\ncorrespondin" +
    "g Office button.\r\n";
            // 
            // InstallForOffice32
            // 
            this.InstallForOffice32.Location = new System.Drawing.Point(12, 117);
            this.InstallForOffice32.Name = "InstallForOffice32";
            this.InstallForOffice32.Size = new System.Drawing.Size(104, 23);
            this.InstallForOffice32.TabIndex = 2;
            this.InstallForOffice32.Text = "Office 32bit";
            this.InstallForOffice32.UseVisualStyleBackColor = true;
            this.InstallForOffice32.Click += new System.EventHandler(this.InstallForOffice32_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(9, 74);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(237, 26);
            this.linkLabel1.TabIndex = 3;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Click here to show Office documentation on\r\nhow to get the bit-version your Offic" +
    "e application.";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // InstallForOffice64
            // 
            this.InstallForOffice64.Location = new System.Drawing.Point(173, 117);
            this.InstallForOffice64.Name = "InstallForOffice64";
            this.InstallForOffice64.Size = new System.Drawing.Size(107, 23);
            this.InstallForOffice64.TabIndex = 4;
            this.InstallForOffice64.Text = "Office 64bit";
            this.InstallForOffice64.UseVisualStyleBackColor = true;
            this.InstallForOffice64.Click += new System.EventHandler(this.InstallForOffice64_Click);
            // 
            // InstallArchSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(289, 152);
            this.Controls.Add(this.InstallForOffice64);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.InstallForOffice32);
            this.Controls.Add(this.label1);
            this.Name = "InstallArchSelector";
            this.ShowIcon = false;
            this.Text = "Unknown Office architecture";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button InstallForOffice32;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button InstallForOffice64;
    }
}