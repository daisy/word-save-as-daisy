

namespace DaisyInstaller {

    partial class PipelineInstaller {
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
            this.installProgress = new System.Windows.Forms.ProgressBar();
            this.Unzipping = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // installProgress
            // 
            this.installProgress.Location = new System.Drawing.Point(11, 32);
            this.installProgress.Name = "installProgress";
            this.installProgress.Size = new System.Drawing.Size(359, 25);
            this.installProgress.TabIndex = 0;
            // 
            // Unzipping
            // 
            this.Unzipping.AutoSize = true;
            this.Unzipping.Location = new System.Drawing.Point(12, 9);
            this.Unzipping.Name = "Unzipping";
            this.Unzipping.Size = new System.Drawing.Size(0, 13);
            this.Unzipping.TabIndex = 1;
            // 
            // PipelineInstaller
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(382, 69);
            this.Controls.Add(this.Unzipping);
            this.Controls.Add(this.installProgress);
            this.Name = "PipelineInstaller";
            this.Text = "Installing DAISY pipeline 1";
            this.Shown += new System.EventHandler(this.PipelineInstaller_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar installProgress;
        private System.Windows.Forms.Label Unzipping;
    }
}