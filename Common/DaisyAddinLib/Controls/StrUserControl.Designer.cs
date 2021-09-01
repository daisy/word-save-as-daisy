namespace Daisy.SaveAsDAISY.Forms.Controls {
    partial class StrUserControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StrUserControl));
            this.m_StrNiceLbl = new System.Windows.Forms.Label();
            this.m_StrTxtBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // m_StrNiceLbl
            // 
            resources.ApplyResources(this.m_StrNiceLbl, "m_StrNiceLbl");
            this.m_StrNiceLbl.Name = "m_StrNiceLbl";
            // 
            // m_StrTxtBox
            // 
            resources.ApplyResources(this.m_StrTxtBox, "m_StrTxtBox");
            this.m_StrTxtBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.m_StrTxtBox.Name = "m_StrTxtBox";
            // 
            // StrUserControl
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.m_StrNiceLbl);
            this.Controls.Add(this.m_StrTxtBox);
            this.Name = "StrUserControl";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label m_StrNiceLbl;
        private System.Windows.Forms.TextBox m_StrTxtBox;
    }
}
