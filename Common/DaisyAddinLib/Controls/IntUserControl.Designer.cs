namespace Daisy.SaveAsDAISY.Forms.Controls
{
    partial class IntUserControl
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
            this.mListBokBox = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.mListBokBox)).BeginInit();
            this.SuspendLayout();
            // 
            // parameterNiceName
            // 
            this.parameterNiceName.Location = new System.Drawing.Point(3, -3);
            // 
            // mListBokBox
            // 
            this.mListBokBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.mListBokBox.Location = new System.Drawing.Point(6, 24);
            this.mListBokBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.mListBokBox.Maximum = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.mListBokBox.Name = "mListBokBox";
            this.mListBokBox.Size = new System.Drawing.Size(362, 23);
            this.mListBokBox.TabIndex = 2;
            // 
            // IntUserControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.mListBokBox);
            this.Margin = new System.Windows.Forms.Padding(5, 7, 5, 7);
            this.Name = "IntUserControl";
            this.Size = new System.Drawing.Size(371, 57);
            this.Controls.SetChildIndex(this.mListBokBox, 0);
            this.Controls.SetChildIndex(this.parameterNiceName, 0);
            ((System.ComponentModel.ISupportInitialize)(this.mListBokBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.NumericUpDown mListBokBox;
    }



}
