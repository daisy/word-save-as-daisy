
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
            this.Message = new System.Windows.Forms.Label();
            this.ExceptionMessage = new System.Windows.Forms.TextBox();
            this.SendReport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Message
            // 
            this.Message.AutoSize = true;
            this.Message.Location = new System.Drawing.Point(13, 14);
            this.Message.Name = "Message";
            this.Message.Size = new System.Drawing.Size(412, 26);
            this.Message.TabIndex = 0;
            this.Message.Text = "An unhandled exception was raised during the execution of the SaveAsDAISY Add-in." +
    "\r\nPlease send the following report to the DAISY pipeline development team:";
            // 
            // ExceptionMessage
            // 
            this.ExceptionMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExceptionMessage.Location = new System.Drawing.Point(12, 43);
            this.ExceptionMessage.Multiline = true;
            this.ExceptionMessage.Name = "ExceptionMessage";
            this.ExceptionMessage.ReadOnly = true;
            this.ExceptionMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.ExceptionMessage.Size = new System.Drawing.Size(413, 369);
            this.ExceptionMessage.TabIndex = 1;
            this.ExceptionMessage.Text = "placeholder";
            // 
            // SendReport
            // 
            this.SendReport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.SendReport.Location = new System.Drawing.Point(348, 418);
            this.SendReport.Name = "SendReport";
            this.SendReport.Size = new System.Drawing.Size(77, 20);
            this.SendReport.TabIndex = 2;
            this.SendReport.Text = "Send report";
            this.SendReport.UseVisualStyleBackColor = true;
            this.SendReport.Click += new System.EventHandler(this.SendReport_Click);
            // 
            // ExceptionReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 450);
            this.Controls.Add(this.SendReport);
            this.Controls.Add(this.ExceptionMessage);
            this.Controls.Add(this.Message);
            this.Name = "ExceptionReport";
            this.ShowIcon = false;
            this.Text = "Unhandled Exception report";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Message;
        private System.Windows.Forms.TextBox ExceptionMessage;
        private System.Windows.Forms.Button SendReport;
    }
}