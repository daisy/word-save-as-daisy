using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion
{
    partial class ConverterSettingsForm
    {
        readonly Dictionary<string, string> positionsMap = new Dictionary<string, string>() {
            { "End of pages", "page"},
            { "Inlined in levels", "inline"}, // TO BE ADDED
            { "End of levels", "end"}

        };

        readonly Dictionary<string, string> levelsMap = new Dictionary<string, string>() {
            //{ "Fifth or nearest parent", "-5" },
            //{ "Fourth or nearest parent", "-4" },
            //{ "Third or nearest parent", "-3" },
            //{ "Second or nearest parent", "-2" },
            //{ "First or nearest parent", "-1" },
            { "Note reference level", "0" },
            { "First or nearest level", "1" },
            { "Second or nearest level", "2" },
            { "Third or nearest level", "3" },
            { "Fourth or nearest level", "4" },
            { "Fifth or nearest level", "5" },
            { "Sixth or nearest level", "6" }
        };

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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.footnotesLevelSelector = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.footnotesPositionSelector = new System.Windows.Forms.ComboBox();
            this.grpbox_pgnum.SuspendLayout();
            this.grpbox_charstyles.SuspendLayout();
            this.grpbox_ImageSizes.SuspendLayout();
            this.groupBox1.SuspendLayout();
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
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.footnotesLevelSelector);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.footnotesPositionSelector);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // footNotesLevel
            // 
            this.footnotesLevelSelector.FormattingEnabled = true;
            this.footnotesLevelSelector.Items.AddRange(new object[] {
            resources.GetString("footNotesLevel.Items"),
            resources.GetString("footNotesLevel.Items1"),
            resources.GetString("footNotesLevel.Items2"),
            resources.GetString("footNotesLevel.Items3"),
            resources.GetString("footNotesLevel.Items4"),
            resources.GetString("footNotesLevel.Items5"),
            resources.GetString("footNotesLevel.Items6")});
            resources.ApplyResources(this.footnotesLevelSelector, "footNotesLevel");
            this.footnotesLevelSelector.Name = "footNotesLevel";
            this.footnotesLevelSelector.SelectedIndexChanged += new System.EventHandler(this.footnotesLevelSelector_SelectedIndexChanged);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // footNotesPosition
            // 
            this.footnotesPositionSelector.FormattingEnabled = true;
            this.footnotesPositionSelector.Items.AddRange(new object[] {
            resources.GetString("footNotesPosition.Items"),
            resources.GetString("footNotesPosition.Items1"),
            resources.GetString("footNotesPosition.Items2")});
            resources.ApplyResources(this.footnotesPositionSelector, "footNotesPosition");
            this.footnotesPositionSelector.Name = "footNotesPosition";
            this.footnotesPositionSelector.SelectedIndexChanged += new System.EventHandler(this.footnotesPositionSelector_SelectedIndexChanged);
            // 
            // ConverterSettingsForm
            // 
            this.AcceptButton = this.btn_ok;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.grpbox_ImageSizes);
            this.Controls.Add(this.grpbox_charstyles);
            this.Controls.Add(this.grpbox_pgnum);
            this.Controls.Add(this.btn_ok);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "ConverterSettingsForm";
            this.ShowIcon = false;
            this.Load += new System.EventHandler(this.Daisysettingsfrm_Load);
            this.grpbox_pgnum.ResumeLayout(false);
            this.grpbox_pgnum.PerformLayout();
            this.grpbox_charstyles.ResumeLayout(false);
            this.grpbox_charstyles.PerformLayout();
            this.grpbox_ImageSizes.ResumeLayout(false);
            this.grpbox_ImageSizes.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
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
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox footnotesPositionSelector;
        private System.Windows.Forms.ComboBox footnotesLevelSelector;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}