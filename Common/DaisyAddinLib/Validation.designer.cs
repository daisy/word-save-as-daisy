namespace Daisy.SaveAsDAISY.Conversion
{
    partial class Validation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Validation));
            this.lbl_Summary = new System.Windows.Forms.Label();
            this.btn_OK = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.linkLabel_Guidelines = new System.Windows.Forms.LinkLabel();
            this.lbl_Information = new System.Windows.Forms.Label();
            this.listBox_Validation = new System.Windows.Forms.ListBox();
            this.saveloglinkLabel = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // lbl_Summary
            // 
            resources.ApplyResources(this.lbl_Summary, "lbl_Summary");
            this.lbl_Summary.Name = "lbl_Summary";
            // 
            // btn_OK
            // 
            this.btn_OK.BackColor = System.Drawing.SystemColors.ControlLight;
            resources.ApplyResources(this.btn_OK, "btn_OK");
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.UseVisualStyleBackColor = false;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btn_Cancel, "btn_Cancel");
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // linkLabel_Guidelines
            // 
            resources.ApplyResources(this.linkLabel_Guidelines, "linkLabel_Guidelines");
            this.linkLabel_Guidelines.LinkColor = System.Drawing.SystemColors.ControlText;
            this.linkLabel_Guidelines.Name = "linkLabel_Guidelines";
            this.linkLabel_Guidelines.TabStop = true;
            this.linkLabel_Guidelines.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lbl_Information
            // 
            this.lbl_Information.BackColor = System.Drawing.SystemColors.Control;
            resources.ApplyResources(this.lbl_Information, "lbl_Information");
            this.lbl_Information.Name = "lbl_Information";
            // 
            // listBox_Validation
            // 
            resources.ApplyResources(this.listBox_Validation, "listBox_Validation");
            this.listBox_Validation.FormattingEnabled = true;
            this.listBox_Validation.Name = "listBox_Validation";
            // 
            // saveloglinkLabel
            // 
            resources.ApplyResources(this.saveloglinkLabel, "saveloglinkLabel");
            this.saveloglinkLabel.LinkColor = System.Drawing.Color.Black;
            this.saveloglinkLabel.Name = "saveloglinkLabel";
            this.saveloglinkLabel.TabStop = true;
            this.saveloglinkLabel.VisitedLinkColor = System.Drawing.Color.Purple;
            this.saveloglinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.saveloglinkLabel_LinkClicked);
            // 
            // Validation
            // 
            this.AcceptButton = this.btn_OK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancel;
            this.Controls.Add(this.saveloglinkLabel);
            this.Controls.Add(this.lbl_Information);
            this.Controls.Add(this.linkLabel_Guidelines);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.lbl_Summary);
            this.Controls.Add(this.listBox_Validation);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Validation";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_Summary;
        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.LinkLabel linkLabel_Guidelines;
        private System.Windows.Forms.Label lbl_Information;
        private System.Windows.Forms.ListBox listBox_Validation;
        private System.Windows.Forms.LinkLabel saveloglinkLabel;
    }
}