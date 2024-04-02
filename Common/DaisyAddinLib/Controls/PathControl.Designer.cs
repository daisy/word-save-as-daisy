namespace Daisy.SaveAsDAISY.Forms.Controls
{
    partial class PathControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.browseButton = new System.Windows.Forms.Button();
            this.parameterValue = new System.Windows.Forms.TextBox();
            this.InputPanel = new System.Windows.Forms.Panel();
            this.InputPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // parameterNiceName
            // 
            this.parameterNiceName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)));
            this.parameterNiceName.Location = new System.Drawing.Point(3, 0);
            this.parameterNiceName.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // browseButton
            // 
            this.browseButton.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.browseButton.Location = new System.Drawing.Point(515, 2);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(73, 27);
            this.browseButton.TabIndex = 3;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = true;
            this.browseButton.Click += new System.EventHandler(this.OnClickBrowseButton);
            // 
            // parameterValue
            // 
            this.parameterValue.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.parameterValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.parameterValue.Location = new System.Drawing.Point(3, 2);
            this.parameterValue.Name = "parameterValue";
            this.parameterValue.Size = new System.Drawing.Size(509, 27);
            this.parameterValue.TabIndex = 2;
            // 
            // InputPanel
            // 
            this.InputPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.InputPanel.Controls.Add(this.parameterValue);
            this.InputPanel.Controls.Add(this.browseButton);
            this.InputPanel.Location = new System.Drawing.Point(3, 23);
            this.InputPanel.Name = "InputPanel";
            this.InputPanel.Size = new System.Drawing.Size(591, 31);
            this.InputPanel.TabIndex = 4;
            // 
            // PathControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.Controls.Add(this.InputPanel);
            this.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.Name = "PathControl";
            this.Size = new System.Drawing.Size(603, 61);
            this.Controls.SetChildIndex(this.parameterNiceName, 0);
            this.Controls.SetChildIndex(this.InputPanel, 0);
            this.InputPanel.ResumeLayout(false);
            this.InputPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.TextBox parameterValue;
        private System.Windows.Forms.Panel InputPanel;
    }
}
