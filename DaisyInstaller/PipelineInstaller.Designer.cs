

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
            this.SuspendLayout();
            // 
            // installProgress
            // 
            this.installProgress.Location = new System.Drawing.Point(11, 12);
            this.installProgress.Name = "installProgress";
            this.installProgress.Size = new System.Drawing.Size(359, 25);
            this.installProgress.TabIndex = 0;
            // 
            // PipelineInstaller
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(382, 54);
            this.Controls.Add(this.installProgress);
            this.Name = "PipelineInstaller";
            this.Text = "Installing pipeline";
            this.Load += new System.EventHandler(this.PipelineInstaller_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar installProgress;
    }
}