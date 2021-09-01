namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    partial class Abbreviation2003
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Abbreviation2003));
            this.lBx_Abbreviation = new System.Windows.Forms.ListBox();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btn_Unmark = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_Goto = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lBx_Abbreviation
            // 
            this.lBx_Abbreviation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lBx_Abbreviation.FormattingEnabled = true;
            resources.ApplyResources(this.lBx_Abbreviation, "lBx_Abbreviation");
            this.lBx_Abbreviation.Name = "lBx_Abbreviation";
            this.lBx_Abbreviation.SelectedValueChanged += new System.EventHandler(this.lBx_Abbreviation_SelectedValueChanged);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btn_Cancel, "btn_Cancel");
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = false;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Unmark
            // 
            this.btn_Unmark.BackColor = System.Drawing.SystemColors.ControlLight;
            resources.ApplyResources(this.btn_Unmark, "btn_Unmark");
            this.btn_Unmark.Name = "btn_Unmark";
            this.btn_Unmark.UseVisualStyleBackColor = false;
            this.btn_Unmark.Click += new System.EventHandler(this.btn_Unmark_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // btn_Goto
            // 
            this.btn_Goto.BackColor = System.Drawing.SystemColors.ControlLight;
            resources.ApplyResources(this.btn_Goto, "btn_Goto");
            this.btn_Goto.Name = "btn_Goto";
            this.btn_Goto.UseVisualStyleBackColor = false;
            this.btn_Goto.Click += new System.EventHandler(this.btn_Goto_Click);
            // 
            // Abbreviation2003
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancel;
            this.Controls.Add(this.btn_Goto);
            this.Controls.Add(this.lBx_Abbreviation);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Unmark);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Abbreviation2003";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lBx_Abbreviation;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Unmark;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_Goto;
    }
}