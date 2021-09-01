namespace Daisy.SaveAsDAISY.Forms.Controls {
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
            this.mIntNiceLabel = new System.Windows.Forms.Label();
            this.mListBokBox = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.mListBokBox)).BeginInit();
            this.SuspendLayout();
            // 
            // mIntNiceLabel
            // 
            this.mIntNiceLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.mIntNiceLabel.AutoSize = true;
            this.mIntNiceLabel.Location = new System.Drawing.Point(94, 2);
            this.mIntNiceLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.mIntNiceLabel.Name = "mIntNiceLabel";
            this.mIntNiceLabel.Size = new System.Drawing.Size(34, 15);
            this.mIntNiceLabel.TabIndex = 1;
            this.mIntNiceLabel.Text = "Nice:";
            // 
            // mListBokBox
            // 
            this.mListBokBox.Location = new System.Drawing.Point(90, 0);
            this.mListBokBox.Maximum = new decimal(new int[] {
            0,
            0,
            0,
            0});
            this.mListBokBox.Name = "mListBokBox";
            this.mListBokBox.Size = new System.Drawing.Size(215, 23);
            this.mListBokBox.TabIndex = 2;
            // 
            // IntUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.mListBokBox);
            this.Controls.Add(this.mIntNiceLabel);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "IntUserControl";
            this.Size = new System.Drawing.Size(361, 26);
            ((System.ComponentModel.ISupportInitialize)(this.mListBokBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label mIntNiceLabel;
        private System.Windows.Forms.NumericUpDown mListBokBox;
    }



}
