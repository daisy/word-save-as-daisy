
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
            this.ConversionProgressBar = new System.Windows.Forms.ProgressBar();
            this.LastMessage = new System.Windows.Forms.Label();
            this.ShowHideDetails = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // MessageTextArea
            // 
            this.MessageTextArea.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MessageTextArea.Location = new System.Drawing.Point(16, 102);
            this.MessageTextArea.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MessageTextArea.Multiline = true;
            this.MessageTextArea.Name = "MessageTextArea";
            this.MessageTextArea.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.MessageTextArea.Size = new System.Drawing.Size(739, 0);
            this.MessageTextArea.TabIndex = 0;
            this.MessageTextArea.Visible = false;
            // 
            // CancelButton
            // 
            this.CancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(659, 116);
            this.CancelButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(100, 28);
            this.CancelButton.TabIndex = 1;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ConversionProgressBar
            // 
            this.ConversionProgressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ConversionProgressBar.Location = new System.Drawing.Point(16, 15);
            this.ConversionProgressBar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ConversionProgressBar.Name = "ConversionProgressBar";
            this.ConversionProgressBar.Size = new System.Drawing.Size(743, 28);
            this.ConversionProgressBar.TabIndex = 2;
            // 
            // LastMessage
            // 
            this.LastMessage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LastMessage.AutoSize = true;
            this.LastMessage.Location = new System.Drawing.Point(16, 47);
            this.LastMessage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.LastMessage.Name = "LastMessage";
            this.LastMessage.Size = new System.Drawing.Size(246, 16);
            this.LastMessage.TabIndex = 3;
            this.LastMessage.Text = "Progression message is displayed here";
            // 
            // ShowHideDetails
            // 
            this.ShowHideDetails.Location = new System.Drawing.Point(16, 66);
            this.ShowHideDetails.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ShowHideDetails.Name = "ShowHideDetails";
            this.ShowHideDetails.Size = new System.Drawing.Size(145, 28);
            this.ShowHideDetails.TabIndex = 4;
            this.ShowHideDetails.Text = "Show details >>";
            this.ShowHideDetails.UseVisualStyleBackColor = true;
            this.ShowHideDetails.Click += new System.EventHandler(this.ShowHideDetails_Click);
            // 
            // ConversionProgress
            // 
            this.AccessibleDescription = "Conversion progress";
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(779, 155);
            this.Controls.Add(this.ShowHideDetails);
            this.Controls.Add(this.LastMessage);
            this.Controls.Add(this.ConversionProgressBar);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.MessageTextArea);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "ConversionProgress";
            this.ShowIcon = false;
            this.Text = "Conversion progress";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ConversionProgress_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox MessageTextArea;
        private new System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.ProgressBar ConversionProgressBar;
        private System.Windows.Forms.Label LastMessage;
        private System.Windows.Forms.Button ShowHideDetails;
    }
}