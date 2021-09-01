namespace Daisy.SaveAsDAISY.Conversion
{
    partial class MasterSubValidation
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MasterSubValidation));
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.OK = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.saveloglinkLabel = new System.Windows.Forms.LinkLabel();
            this.linkLabel_Guidelines = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btn_Cancel, "btn_Cancel");
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = false;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // OK
            // 
            this.OK.BackColor = System.Drawing.SystemColors.ControlLight;
            this.OK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.OK, "OK");
            this.OK.Name = "OK";
            this.OK.UseVisualStyleBackColor = false;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // richTextBox1
            // 
            resources.ApplyResources(this.richTextBox1, "richTextBox1");
            this.richTextBox1.Name = "richTextBox1";
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
            // linkLabel_Guidelines
            // 
            resources.ApplyResources(this.linkLabel_Guidelines, "linkLabel_Guidelines");
            this.linkLabel_Guidelines.LinkColor = System.Drawing.SystemColors.ControlText;
            this.linkLabel_Guidelines.Name = "linkLabel_Guidelines";
            this.linkLabel_Guidelines.TabStop = true;
            this.linkLabel_Guidelines.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_Guidelines_LinkClicked);
            // 
            // MasterSubValidation
            // 
            this.AcceptButton = this.OK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancel;
            this.Controls.Add(this.saveloglinkLabel);
            this.Controls.Add(this.linkLabel_Guidelines);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.OK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MasterSubValidation";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.LinkLabel saveloglinkLabel;
        private System.Windows.Forms.LinkLabel linkLabel_Guidelines;
    }
}