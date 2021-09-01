namespace Daisy.SaveAsDAISY.Forms.Controls {
    partial class PathBrowserControl
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PathBrowserControl));
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.mNiceNameLabel = new System.Windows.Forms.Label();
            this.mTextBox = new System.Windows.Forms.TextBox();
            this.mBrowseButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // tableLayoutPanel1
            // 
            resources.ApplyResources(this.tableLayoutPanel1, "tableLayoutPanel1");
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel1.Controls.Add(this.mNiceNameLabel, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.mTextBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.mBrowseButton, 2, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            // 
            // mNiceNameLabel
            // 
            resources.ApplyResources(this.mNiceNameLabel, "mNiceNameLabel");
            this.mNiceNameLabel.Name = "mNiceNameLabel";
            // 
            // mTextBox
            // 
            resources.ApplyResources(this.mTextBox, "mTextBox");
            this.mTextBox.Name = "mTextBox";
            // 
            // mBrowseButton
            // 
            resources.ApplyResources(this.mBrowseButton, "mBrowseButton");
            this.mBrowseButton.Name = "mBrowseButton";
            this.mBrowseButton.UseVisualStyleBackColor = true;
            this.mBrowseButton.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // PathBrowserControl
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "PathBrowserControl";
            this.Load += new System.EventHandler(this.PathBrowserControl_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label mNiceNameLabel;
        private System.Windows.Forms.Button mBrowseButton;
        private System.Windows.Forms.TextBox mTextBox;
    }

}
