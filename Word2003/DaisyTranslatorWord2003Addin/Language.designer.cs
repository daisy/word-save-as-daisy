namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    partial class Language
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Language));
            this.lBx_Lngs = new System.Windows.Forms.ListBox();
            this.lbl_Lang = new System.Windows.Forms.Label();
            this.btn_Apply = new System.Windows.Forms.Button();
            this.btn_Close = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lBx_Lngs
            // 
            this.lBx_Lngs.FormattingEnabled = true;
            resources.ApplyResources(this.lBx_Lngs, "lBx_Lngs");
            this.lBx_Lngs.Name = "lBx_Lngs";
            // 
            // lbl_Lang
            // 
            resources.ApplyResources(this.lbl_Lang, "lbl_Lang");
            this.lbl_Lang.Name = "lbl_Lang";
            // 
            // btn_Apply
            // 
            resources.ApplyResources(this.btn_Apply, "btn_Apply");
            this.btn_Apply.Name = "btn_Apply";
            this.btn_Apply.UseVisualStyleBackColor = true;
            this.btn_Apply.Click += new System.EventHandler(this.btn_Apply_Click);
            // 
            // btn_Close
            // 
            this.btn_Close.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btn_Close, "btn_Close");
            this.btn_Close.Name = "btn_Close";
            this.btn_Close.UseVisualStyleBackColor = true;
            this.btn_Close.Click += new System.EventHandler(this.btn_Close_Click);
            // 
            // Language
            // 
            this.AcceptButton = this.btn_Apply;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Close;
            this.Controls.Add(this.btn_Close);
            this.Controls.Add(this.btn_Apply);
            this.Controls.Add(this.lbl_Lang);
            this.Controls.Add(this.lBx_Lngs);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Language";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lBx_Lngs;
        private System.Windows.Forms.Label lbl_Lang;
        private System.Windows.Forms.Button btn_Apply;
        private System.Windows.Forms.Button btn_Close;
    }
}