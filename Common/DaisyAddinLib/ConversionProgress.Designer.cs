
namespace Daisy.SaveAsDAISY {
    partial class ConversionProgress {
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
            this.MessageTextArea = new System.Windows.Forms.TextBox();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // MessageTextArea
            // 
            this.MessageTextArea.Location = new System.Drawing.Point(12, 12);
            this.MessageTextArea.Multiline = true;
            this.MessageTextArea.Name = "MessageTextArea";
            this.MessageTextArea.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MessageTextArea.Size = new System.Drawing.Size(617, 391);
            this.MessageTextArea.TabIndex = 0;
            // 
            // CancelButton
            // 
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(554, 415);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(75, 23);
            this.CancelButton.TabIndex = 1;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ConversionProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 450);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.MessageTextArea);
            this.Name = "ConversionProgress";
            this.ShowIcon = false;
            this.Text = "Convertion to DTBook XML";
            this.ResumeLayout(false);
            this.PerformLayout();
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(ConversionProgress_FormClosing);

        }

        #endregion

        private System.Windows.Forms.TextBox MessageTextArea;
        private new System.Windows.Forms.Button CancelButton;
    }
}