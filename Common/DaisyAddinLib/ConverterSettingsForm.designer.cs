namespace Daisy.SaveAsDAISY.Conversion
{
    partial class ConverterSettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConverterSettingsForm));
            this.radiobtn_custom = new System.Windows.Forms.RadioButton();
            this.radiobtn_auto = new System.Windows.Forms.RadioButton();
            this.checkbox_translate = new System.Windows.Forms.CheckBox();
            this.radiobtn_originalimage = new System.Windows.Forms.RadioButton();
            this.radiobtn_resize = new System.Windows.Forms.RadioButton();
            this.radiobtn_resample = new System.Windows.Forms.RadioButton();
            this.combobox_resample = new System.Windows.Forms.ComboBox();
            this.btn_ok = new System.Windows.Forms.Button();
            this.grpbox_pgnum = new System.Windows.Forms.GroupBox();
            this.grpbox_charstyles = new System.Windows.Forms.GroupBox();
            this.grpbox_ImageSizes = new System.Windows.Forms.GroupBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.grpbox_pgnum.SuspendLayout();
            this.grpbox_charstyles.SuspendLayout();
            this.grpbox_ImageSizes.SuspendLayout();
            this.SuspendLayout();
            // 
            // radiobtn_custom
            // 
            resources.ApplyResources(this.radiobtn_custom, "radiobtn_custom");
            this.radiobtn_custom.Name = "radiobtn_custom";
            this.radiobtn_custom.TabStop = true;
            this.radiobtn_custom.UseVisualStyleBackColor = true;
            // 
            // radiobtn_auto
            // 
            resources.ApplyResources(this.radiobtn_auto, "radiobtn_auto");
            this.radiobtn_auto.Name = "radiobtn_auto";
            this.radiobtn_auto.TabStop = true;
            this.radiobtn_auto.UseVisualStyleBackColor = true;
            // 
            // checkbox_translate
            // 
            resources.ApplyResources(this.checkbox_translate, "checkbox_translate");
            this.checkbox_translate.Name = "checkbox_translate";
            this.checkbox_translate.UseVisualStyleBackColor = true;
            // 
            // radiobtn_originalimage
            // 
            resources.ApplyResources(this.radiobtn_originalimage, "radiobtn_originalimage");
            this.radiobtn_originalimage.Name = "radiobtn_originalimage";
            this.radiobtn_originalimage.TabStop = true;
            this.radiobtn_originalimage.UseVisualStyleBackColor = true;
            this.radiobtn_originalimage.Click += new System.EventHandler(this.radiobtn_originalimage_Click);
            // 
            // radiobtn_resize
            // 
            resources.ApplyResources(this.radiobtn_resize, "radiobtn_resize");
            this.radiobtn_resize.Name = "radiobtn_resize";
            this.radiobtn_resize.TabStop = true;
            this.radiobtn_resize.UseVisualStyleBackColor = true;
            this.radiobtn_resize.Click += new System.EventHandler(this.radiobtn_resize_Click);
            // 
            // radiobtn_resample
            // 
            resources.ApplyResources(this.radiobtn_resample, "radiobtn_resample");
            this.radiobtn_resample.Name = "radiobtn_resample";
            this.radiobtn_resample.TabStop = true;
            this.radiobtn_resample.UseVisualStyleBackColor = true;
            this.radiobtn_resample.Click += new System.EventHandler(this.radiobtn_resample_Click);
            this.radiobtn_resample.MouseClick += new System.Windows.Forms.MouseEventHandler(this.radiobtn_resample_MouseClick);
            // 
            // combobox_resample
            // 
            this.combobox_resample.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.combobox_resample.FormattingEnabled = true;
            this.combobox_resample.Items.AddRange(new object[] {
            resources.GetString("combobox_resample.Items"),
            resources.GetString("combobox_resample.Items1"),
            resources.GetString("combobox_resample.Items2")});
            resources.ApplyResources(this.combobox_resample, "combobox_resample");
            this.combobox_resample.Name = "combobox_resample";
            // 
            // btn_ok
            // 
            resources.ApplyResources(this.btn_ok, "btn_ok");
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // grpbox_pgnum
            // 
            this.grpbox_pgnum.Controls.Add(this.radiobtn_auto);
            this.grpbox_pgnum.Controls.Add(this.radiobtn_custom);
            resources.ApplyResources(this.grpbox_pgnum, "grpbox_pgnum");
            this.grpbox_pgnum.Name = "grpbox_pgnum";
            this.grpbox_pgnum.TabStop = false;
            // 
            // grpbox_charstyles
            // 
            this.grpbox_charstyles.Controls.Add(this.checkbox_translate);
            resources.ApplyResources(this.grpbox_charstyles, "grpbox_charstyles");
            this.grpbox_charstyles.Name = "grpbox_charstyles";
            this.grpbox_charstyles.TabStop = false;
            // 
            // grpbox_ImageSizes
            // 
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_originalimage);
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_resize);
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_resample);
            this.grpbox_ImageSizes.Controls.Add(this.combobox_resample);
            resources.ApplyResources(this.grpbox_ImageSizes, "grpbox_ImageSizes");
            this.grpbox_ImageSizes.Name = "grpbox_ImageSizes";
            this.grpbox_ImageSizes.TabStop = false;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.button1_Click);
            // 
            // DAISY_Settings
            // 
            this.AcceptButton = this.btn_ok;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.grpbox_ImageSizes);
            this.Controls.Add(this.grpbox_charstyles);
            this.Controls.Add(this.grpbox_pgnum);
            this.Controls.Add(this.btn_ok);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "DAISY_Settings";
            this.ShowIcon = false;
            this.Load += new System.EventHandler(this.Daisysettingsfrm_Load);
            this.grpbox_pgnum.ResumeLayout(false);
            this.grpbox_pgnum.PerformLayout();
            this.grpbox_charstyles.ResumeLayout(false);
            this.grpbox_charstyles.PerformLayout();
            this.grpbox_ImageSizes.ResumeLayout(false);
            this.grpbox_ImageSizes.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radiobtn_custom;
        private System.Windows.Forms.RadioButton radiobtn_auto;
        private System.Windows.Forms.CheckBox checkbox_translate;
        private System.Windows.Forms.RadioButton radiobtn_originalimage;
        private System.Windows.Forms.RadioButton radiobtn_resize;
        private System.Windows.Forms.RadioButton radiobtn_resample;
        private System.Windows.Forms.ComboBox combobox_resample;
        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.GroupBox grpbox_pgnum;
        private System.Windows.Forms.GroupBox grpbox_charstyles;
        private System.Windows.Forms.GroupBox grpbox_ImageSizes;
        private System.Windows.Forms.Button btnCancel;
    }
}