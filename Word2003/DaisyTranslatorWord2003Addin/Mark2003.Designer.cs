namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    partial class Mark2003
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mark2003));
            this.cBx_AllOccurences = new System.Windows.Forms.CheckBox();
            this.cBx_Pronounce = new System.Windows.Forms.CheckBox();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Mark = new System.Windows.Forms.Button();
            this.tBx_MarkFullForm = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox_Settings = new System.Windows.Forms.GroupBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox_Settings.SuspendLayout();
            this.SuspendLayout();
            // 
            // cBx_AllOccurences
            // 
            resources.ApplyResources(this.cBx_AllOccurences, "cBx_AllOccurences");
            this.cBx_AllOccurences.Name = "cBx_AllOccurences";
            this.toolTip1.SetToolTip(this.cBx_AllOccurences, resources.GetString("cBx_AllOccurences.ToolTip"));
            this.cBx_AllOccurences.UseVisualStyleBackColor = true;
            this.cBx_AllOccurences.CheckedChanged += new System.EventHandler(this.cBx_AllOccurences_CheckedChanged);
            // 
            // cBx_Pronounce
            // 
            resources.ApplyResources(this.cBx_Pronounce, "cBx_Pronounce");
            this.cBx_Pronounce.Name = "cBx_Pronounce";
            this.toolTip1.SetToolTip(this.cBx_Pronounce, resources.GetString("cBx_Pronounce.ToolTip"));
            this.cBx_Pronounce.UseVisualStyleBackColor = true;
            this.cBx_Pronounce.CheckedChanged += new System.EventHandler(this.cBx_Pronounce_CheckedChanged);
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
            // btn_Mark
            // 
            this.btn_Mark.BackColor = System.Drawing.SystemColors.ControlLight;
            resources.ApplyResources(this.btn_Mark, "btn_Mark");
            this.btn_Mark.Name = "btn_Mark";
            this.btn_Mark.UseVisualStyleBackColor = false;
            this.btn_Mark.Click += new System.EventHandler(this.btn_Mark_Click);
            // 
            // tBx_MarkFullForm
            // 
            this.tBx_MarkFullForm.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.tBx_MarkFullForm, "tBx_MarkFullForm");
            this.tBx_MarkFullForm.Name = "tBx_MarkFullForm";
            this.toolTip1.SetToolTip(this.tBx_MarkFullForm, resources.GetString("tBx_MarkFullForm.ToolTip"));
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // groupBox_Settings
            // 
            this.groupBox_Settings.Controls.Add(this.cBx_AllOccurences);
            this.groupBox_Settings.Controls.Add(this.cBx_Pronounce);
            resources.ApplyResources(this.groupBox_Settings, "groupBox_Settings");
            this.groupBox_Settings.Name = "groupBox_Settings";
            this.groupBox_Settings.TabStop = false;
            // 
            // Mark2003
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancel;
            this.Controls.Add(this.groupBox_Settings);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Mark);
            this.Controls.Add(this.tBx_MarkFullForm);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Mark2003";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.groupBox_Settings.ResumeLayout(false);
            this.groupBox_Settings.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox cBx_AllOccurences;
        private System.Windows.Forms.CheckBox cBx_Pronounce;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Mark;
        private System.Windows.Forms.TextBox tBx_MarkFullForm;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox_Settings;
        private System.Windows.Forms.ToolTip toolTip1;


    }
}