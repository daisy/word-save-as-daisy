
namespace Daisy.SaveAsDAISY.Addins.Word2007 {
    partial class ExceptionReport {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExceptionReport));
            this.ExceptionMessage = new System.Windows.Forms.TextBox();
            this.SendReport = new System.Windows.Forms.Button();
            this.Message = new System.Windows.Forms.LinkLabel();
            this.CheckSimilarIssue = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ExceptionMessage
            // 
            this.ExceptionMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExceptionMessage.Location = new System.Drawing.Point(16, 93);
            this.ExceptionMessage.Margin = new System.Windows.Forms.Padding(4);
            this.ExceptionMessage.Multiline = true;
            this.ExceptionMessage.Name = "ExceptionMessage";
            this.ExceptionMessage.ReadOnly = true;
            this.ExceptionMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.ExceptionMessage.Size = new System.Drawing.Size(505, 416);
            this.ExceptionMessage.TabIndex = 1;
            this.ExceptionMessage.Text = "placeholder";
            // 
            // SendReport
            // 
            this.SendReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.SendReport.Location = new System.Drawing.Point(16, 516);
            this.SendReport.Margin = new System.Windows.Forms.Padding(4);
            this.SendReport.Name = "SendReport";
            this.SendReport.Size = new System.Drawing.Size(177, 25);
            this.SendReport.TabIndex = 2;
            this.SendReport.Text = "Report issue on Github";
            this.SendReport.UseVisualStyleBackColor = true;
            this.SendReport.Click += new System.EventHandler(this.SendReport_Click);
            // 
            // Message
            // 
            this.Message.AutoSize = true;
            this.Message.LinkArea = new System.Windows.Forms.LinkArea(168, 32);
            this.Message.Location = new System.Drawing.Point(10, 9);
            this.Message.Name = "Message";
            this.Message.Size = new System.Drawing.Size(621, 80);
            this.Message.TabIndex = 3;
            this.Message.TabStop = true;
            this.Message.Text = resources.GetString("Message.Text");
            this.Message.UseCompatibleTextRendering = true;
            this.Message.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.Message_LinkClicked);
            // 
            // CheckSimilarIssue
            // 
            this.CheckSimilarIssue.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.CheckSimilarIssue.Location = new System.Drawing.Point(278, 516);
            this.CheckSimilarIssue.Name = "CheckSimilarIssue";
            this.CheckSimilarIssue.Size = new System.Drawing.Size(243, 25);
            this.CheckSimilarIssue.TabIndex = 4;
            this.CheckSimilarIssue.Text = "Check for similar issues on GitHub";
            this.CheckSimilarIssue.UseVisualStyleBackColor = true;
            this.CheckSimilarIssue.Click += new System.EventHandler(this.CheckForSimilarIssues);
            // 
            // ExceptionReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 554);
            this.Controls.Add(this.CheckSimilarIssue);
            this.Controls.Add(this.Message);
            this.Controls.Add(this.SendReport);
            this.Controls.Add(this.ExceptionMessage);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ExceptionReport";
            this.ShowIcon = false;
            this.Text = "Exception to report";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox ExceptionMessage;
        private System.Windows.Forms.Button SendReport;
        private System.Windows.Forms.LinkLabel Message;
        private System.Windows.Forms.Button CheckSimilarIssue;
    }
}