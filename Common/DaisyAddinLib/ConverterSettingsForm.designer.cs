using System.Collections.Generic;
using System.Linq;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.Conversion
{
    partial class ConverterSettingsForm
    {
        readonly Dictionary<string, FootnotesPositionChoice.Enum> positionsMap = new Dictionary<string, FootnotesPositionChoice.Enum>() {
            { "End of pages", FootnotesPositionChoice.Enum.Page},
            { "Inlined in levels", FootnotesPositionChoice.Enum.Inline},
            { "End of levels", FootnotesPositionChoice.Enum.End},
        };

        readonly Dictionary<string, int> levelsMap = new Dictionary<string, int>() {
            //{ "Fifth or nearest parent", "-5" },
            //{ "Fourth or nearest parent", "-4" },
            //{ "Third or nearest parent", "-3" },
            //{ "Second or nearest parent", "-2" },
            //{ "First or nearest parent", "-1" },
            { "Note reference level", 0 },
            { "First or nearest level", 1 },
            { "Second or nearest level", 2 },
            { "Third or nearest level", 3 },
            { "Fourth or nearest level", 4 },
            { "Fifth or nearest level", 5 },
            { "Sixth or nearest level", 6 }
        };


        readonly Dictionary<string, FootnotesNumberingChoice.Enum> notesNumberingMap = new Dictionary<string, FootnotesNumberingChoice.Enum>() {
            { "Don't add a numeric prefix", FootnotesNumberingChoice.Enum.None},
            { "Recompute a numeric prefix", FootnotesNumberingChoice.Enum.Number}, 
           // { "Use word computed note number", FootnotesNumberingChoice.Enum.Word}// TO BE ADDED
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
            this.FootnotesBox = new System.Windows.Forms.GroupBox();
            this.notesPositionLabel = new System.Windows.Forms.Label();
            this.notesPositionSelector = new System.Windows.Forms.ComboBox();
            this.notesLevelLabel = new System.Windows.Forms.Label();
            this.notesLevelSelector = new System.Windows.Forms.ComboBox();
            this.notesNumberingLabel = new System.Windows.Forms.Label();
            this.notesNumberingSelector = new System.Windows.Forms.ComboBox();
            this.notesNumberingStartLabel = new System.Windows.Forms.Label();
            this.notesNumberingStartValue = new System.Windows.Forms.MaskedTextBox();
            this.notesNumberPrefixLabel = new System.Windows.Forms.Label();
            this.notesNumberPrefixValue = new System.Windows.Forms.TextBox();
            this.notesNumberSuffixLabel = new System.Windows.Forms.Label();
            this.notesNumberSuffixValue = new System.Windows.Forms.TextBox();
            this.TextToSpeechBox = new System.Windows.Forms.GroupBox();
            this.TTSConfigFileLabel = new System.Windows.Forms.Label();
            this.TTSConfigFilePath = new System.Windows.Forms.TextBox();
            this.TTSConfigFileBrowse = new System.Windows.Forms.Button();
            this.AzureRegionLabel = new System.Windows.Forms.Label();
            this.AzureRegionValue = new System.Windows.Forms.ComboBox();
            this.AzureKeyLabel = new System.Windows.Forms.Label();
            this.AzureKeyValue = new System.Windows.Forms.TextBox();
            this.grpbox_pgnum.SuspendLayout();
            this.grpbox_charstyles.SuspendLayout();
            this.grpbox_ImageSizes.SuspendLayout();
            this.FootnotesBox.SuspendLayout();
            this.TextToSpeechBox.SuspendLayout();
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
            resources.ApplyResources(this.grpbox_pgnum, "grpbox_pgnum");
            this.grpbox_pgnum.Controls.Add(this.radiobtn_auto);
            this.grpbox_pgnum.Controls.Add(this.radiobtn_custom);
            this.grpbox_pgnum.Name = "grpbox_pgnum";
            this.grpbox_pgnum.TabStop = false;
            // 
            // grpbox_charstyles
            // 
            resources.ApplyResources(this.grpbox_charstyles, "grpbox_charstyles");
            this.grpbox_charstyles.Controls.Add(this.checkbox_translate);
            this.grpbox_charstyles.Name = "grpbox_charstyles";
            this.grpbox_charstyles.TabStop = false;
            // 
            // grpbox_ImageSizes
            // 
            resources.ApplyResources(this.grpbox_ImageSizes, "grpbox_ImageSizes");
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_originalimage);
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_resize);
            this.grpbox_ImageSizes.Controls.Add(this.radiobtn_resample);
            this.grpbox_ImageSizes.Controls.Add(this.combobox_resample);
            this.grpbox_ImageSizes.Name = "grpbox_ImageSizes";
            this.grpbox_ImageSizes.TabStop = false;
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.button1_Click);
            // 
            // FootnotesBox
            // 
            resources.ApplyResources(this.FootnotesBox, "FootnotesBox");
            this.FootnotesBox.Controls.Add(this.notesPositionLabel);
            this.FootnotesBox.Controls.Add(this.notesPositionSelector);
            this.FootnotesBox.Controls.Add(this.notesLevelLabel);
            this.FootnotesBox.Controls.Add(this.notesLevelSelector);
            this.FootnotesBox.Controls.Add(this.notesNumberingLabel);
            this.FootnotesBox.Controls.Add(this.notesNumberingSelector);
            this.FootnotesBox.Controls.Add(this.notesNumberingStartLabel);
            this.FootnotesBox.Controls.Add(this.notesNumberingStartValue);
            this.FootnotesBox.Controls.Add(this.notesNumberPrefixLabel);
            this.FootnotesBox.Controls.Add(this.notesNumberPrefixValue);
            this.FootnotesBox.Controls.Add(this.notesNumberSuffixLabel);
            this.FootnotesBox.Controls.Add(this.notesNumberSuffixValue);
            this.FootnotesBox.Name = "FootnotesBox";
            this.FootnotesBox.TabStop = false;
            // 
            // notesPositionLabel
            // 
            resources.ApplyResources(this.notesPositionLabel, "notesPositionLabel");
            this.notesPositionLabel.Name = "notesPositionLabel";
            // 
            // notesPositionSelector
            // 
            resources.ApplyResources(this.notesPositionSelector, "notesPositionSelector");
            this.notesPositionSelector.FormattingEnabled = true;
            this.notesPositionSelector.Items.AddRange(new object[] {
            resources.GetString("notesPositionSelector.Items"),
            resources.GetString("notesPositionSelector.Items1"),
            resources.GetString("notesPositionSelector.Items2")});
            this.notesPositionSelector.Name = "notesPositionSelector";
            this.notesPositionSelector.SelectedIndexChanged += new System.EventHandler(this.footnotesPositionSelector_SelectedIndexChanged);
            // 
            // notesLevelLabel
            // 
            resources.ApplyResources(this.notesLevelLabel, "notesLevelLabel");
            this.notesLevelLabel.Name = "notesLevelLabel";
            // 
            // notesLevelSelector
            // 
            resources.ApplyResources(this.notesLevelSelector, "notesLevelSelector");
            this.notesLevelSelector.FormattingEnabled = true;
            this.notesLevelSelector.Items.AddRange(new object[] {
            resources.GetString("notesLevelSelector.Items"),
            resources.GetString("notesLevelSelector.Items1"),
            resources.GetString("notesLevelSelector.Items2"),
            resources.GetString("notesLevelSelector.Items3"),
            resources.GetString("notesLevelSelector.Items4"),
            resources.GetString("notesLevelSelector.Items5"),
            resources.GetString("notesLevelSelector.Items6")});
            this.notesLevelSelector.Name = "notesLevelSelector";
            this.notesLevelSelector.SelectedIndexChanged += new System.EventHandler(this.footnotesLevelSelector_SelectedIndexChanged);
            // 
            // notesNumberingLabel
            // 
            resources.ApplyResources(this.notesNumberingLabel, "notesNumberingLabel");
            this.notesNumberingLabel.Name = "notesNumberingLabel";
            // 
            // notesNumberingSelector
            // 
            resources.ApplyResources(this.notesNumberingSelector, "notesNumberingSelector");
            this.notesNumberingSelector.FormattingEnabled = true;
            this.notesNumberingSelector.Items.AddRange(new object[] {
            resources.GetString("notesNumberingSelector.Items"),
            resources.GetString("notesNumberingSelector.Items1")});
            this.notesNumberingSelector.Name = "notesNumberingSelector";
            this.notesNumberingSelector.SelectedIndexChanged += new System.EventHandler(this.notesNumberingSelector_SelectedIndexChanged);
            // 
            // notesNumberingStartLabel
            // 
            resources.ApplyResources(this.notesNumberingStartLabel, "notesNumberingStartLabel");
            this.notesNumberingStartLabel.Name = "notesNumberingStartLabel";
            // 
            // notesNumberingStartValue
            // 
            resources.ApplyResources(this.notesNumberingStartValue, "notesNumberingStartValue");
            this.notesNumberingStartValue.Name = "notesNumberingStartValue";
            // 
            // notesNumberPrefixLabel
            // 
            resources.ApplyResources(this.notesNumberPrefixLabel, "notesNumberPrefixLabel");
            this.notesNumberPrefixLabel.Name = "notesNumberPrefixLabel";
            // 
            // notesNumberPrefixValue
            // 
            resources.ApplyResources(this.notesNumberPrefixValue, "notesNumberPrefixValue");
            this.notesNumberPrefixValue.Name = "notesNumberPrefixValue";
            // 
            // notesNumberSuffixLabel
            // 
            resources.ApplyResources(this.notesNumberSuffixLabel, "notesNumberSuffixLabel");
            this.notesNumberSuffixLabel.Name = "notesNumberSuffixLabel";
            // 
            // notesNumberSuffixValue
            // 
            resources.ApplyResources(this.notesNumberSuffixValue, "notesNumberSuffixValue");
            this.notesNumberSuffixValue.Name = "notesNumberSuffixValue";
            // 
            // TextToSpeechBox
            // 
            resources.ApplyResources(this.TextToSpeechBox, "TextToSpeechBox");
            this.TextToSpeechBox.Controls.Add(this.AzureKeyValue);
            this.TextToSpeechBox.Controls.Add(this.AzureKeyLabel);
            this.TextToSpeechBox.Controls.Add(this.AzureRegionValue);
            this.TextToSpeechBox.Controls.Add(this.AzureRegionLabel);
            this.TextToSpeechBox.Controls.Add(this.TTSConfigFileBrowse);
            this.TextToSpeechBox.Controls.Add(this.TTSConfigFilePath);
            this.TextToSpeechBox.Controls.Add(this.TTSConfigFileLabel);
            this.TextToSpeechBox.Name = "TextToSpeechBox";
            this.TextToSpeechBox.TabStop = false;
            // 
            // TTSConfigFileLabel
            // 
            resources.ApplyResources(this.TTSConfigFileLabel, "TTSConfigFileLabel");
            this.TTSConfigFileLabel.Name = "TTSConfigFileLabel";
            // 
            // TTSConfigFilePath
            // 
            resources.ApplyResources(this.TTSConfigFilePath, "TTSConfigFilePath");
            this.TTSConfigFilePath.Name = "TTSConfigFilePath";
            // 
            // TTSConfigFileBrowse
            // 
            resources.ApplyResources(this.TTSConfigFileBrowse, "TTSConfigFileBrowse");
            this.TTSConfigFileBrowse.Name = "TTSConfigFileBrowse";
            this.TTSConfigFileBrowse.UseVisualStyleBackColor = true;
            // 
            // AzureRegionLabel
            // 
            resources.ApplyResources(this.AzureRegionLabel, "AzureRegionLabel");
            this.AzureRegionLabel.Name = "AzureRegionLabel";
            // 
            // AzureRegionValue
            // 
            resources.ApplyResources(this.AzureRegionValue, "AzureRegionValue");
            this.AzureRegionValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.AzureRegionValue.FormattingEnabled = true;
            this.AzureRegionValue.Items.AddRange(new object[] {
            resources.GetString("AzureRegionValue.Items"),
            resources.GetString("AzureRegionValue.Items1"),
            resources.GetString("AzureRegionValue.Items2"),
            resources.GetString("AzureRegionValue.Items3"),
            resources.GetString("AzureRegionValue.Items4"),
            resources.GetString("AzureRegionValue.Items5"),
            resources.GetString("AzureRegionValue.Items6"),
            resources.GetString("AzureRegionValue.Items7"),
            resources.GetString("AzureRegionValue.Items8"),
            resources.GetString("AzureRegionValue.Items9"),
            resources.GetString("AzureRegionValue.Items10"),
            resources.GetString("AzureRegionValue.Items11"),
            resources.GetString("AzureRegionValue.Items12"),
            resources.GetString("AzureRegionValue.Items13"),
            resources.GetString("AzureRegionValue.Items14"),
            resources.GetString("AzureRegionValue.Items15"),
            resources.GetString("AzureRegionValue.Items16"),
            resources.GetString("AzureRegionValue.Items17"),
            resources.GetString("AzureRegionValue.Items18"),
            resources.GetString("AzureRegionValue.Items19"),
            resources.GetString("AzureRegionValue.Items20"),
            resources.GetString("AzureRegionValue.Items21"),
            resources.GetString("AzureRegionValue.Items22"),
            resources.GetString("AzureRegionValue.Items23"),
            resources.GetString("AzureRegionValue.Items24"),
            resources.GetString("AzureRegionValue.Items25"),
            resources.GetString("AzureRegionValue.Items26"),
            resources.GetString("AzureRegionValue.Items27"),
            resources.GetString("AzureRegionValue.Items28"),
            resources.GetString("AzureRegionValue.Items29"),
            resources.GetString("AzureRegionValue.Items30"),
            resources.GetString("AzureRegionValue.Items31"),
            resources.GetString("AzureRegionValue.Items32"),
            resources.GetString("AzureRegionValue.Items33"),
            resources.GetString("AzureRegionValue.Items34"),
            resources.GetString("AzureRegionValue.Items35"),
            resources.GetString("AzureRegionValue.Items36"),
            resources.GetString("AzureRegionValue.Items37"),
            resources.GetString("AzureRegionValue.Items38"),
            resources.GetString("AzureRegionValue.Items39"),
            resources.GetString("AzureRegionValue.Items40"),
            resources.GetString("AzureRegionValue.Items41"),
            resources.GetString("AzureRegionValue.Items42"),
            resources.GetString("AzureRegionValue.Items43"),
            resources.GetString("AzureRegionValue.Items44"),
            resources.GetString("AzureRegionValue.Items45"),
            resources.GetString("AzureRegionValue.Items46"),
            resources.GetString("AzureRegionValue.Items47"),
            resources.GetString("AzureRegionValue.Items48"),
            resources.GetString("AzureRegionValue.Items49"),
            resources.GetString("AzureRegionValue.Items50"),
            resources.GetString("AzureRegionValue.Items51"),
            resources.GetString("AzureRegionValue.Items52"),
            resources.GetString("AzureRegionValue.Items53"),
            resources.GetString("AzureRegionValue.Items54"),
            resources.GetString("AzureRegionValue.Items55"),
            resources.GetString("AzureRegionValue.Items56"),
            resources.GetString("AzureRegionValue.Items57"),
            resources.GetString("AzureRegionValue.Items58"),
            resources.GetString("AzureRegionValue.Items59"),
            resources.GetString("AzureRegionValue.Items60"),
            resources.GetString("AzureRegionValue.Items61"),
            resources.GetString("AzureRegionValue.Items62"),
            resources.GetString("AzureRegionValue.Items63"),
            resources.GetString("AzureRegionValue.Items64"),
            resources.GetString("AzureRegionValue.Items65"),
            resources.GetString("AzureRegionValue.Items66"),
            resources.GetString("AzureRegionValue.Items67"),
            resources.GetString("AzureRegionValue.Items68"),
            resources.GetString("AzureRegionValue.Items69"),
            resources.GetString("AzureRegionValue.Items70"),
            resources.GetString("AzureRegionValue.Items71"),
            resources.GetString("AzureRegionValue.Items72"),
            resources.GetString("AzureRegionValue.Items73"),
            resources.GetString("AzureRegionValue.Items74"),
            resources.GetString("AzureRegionValue.Items75"),
            resources.GetString("AzureRegionValue.Items76"),
            resources.GetString("AzureRegionValue.Items77"),
            resources.GetString("AzureRegionValue.Items78"),
            resources.GetString("AzureRegionValue.Items79"),
            resources.GetString("AzureRegionValue.Items80")});
            this.AzureRegionValue.Name = "AzureRegionValue";
            // 
            // AzureKeyLabel
            // 
            resources.ApplyResources(this.AzureKeyLabel, "AzureKeyLabel");
            this.AzureKeyLabel.Name = "AzureKeyLabel";
            // 
            // AzureKeyValue
            // 
            resources.ApplyResources(this.AzureKeyValue, "AzureKeyValue");
            this.AzureKeyValue.Name = "AzureKeyValue";
            // 
            // ConverterSettingsForm
            // 
            this.AcceptButton = this.btn_ok;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.FootnotesBox);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.grpbox_ImageSizes);
            this.Controls.Add(this.grpbox_charstyles);
            this.Controls.Add(this.grpbox_pgnum);
            this.Controls.Add(this.TextToSpeechBox);
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
            this.FootnotesBox.ResumeLayout(false);
            this.FootnotesBox.PerformLayout();
            this.TextToSpeechBox.ResumeLayout(false);
            this.TextToSpeechBox.PerformLayout();
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
        private System.Windows.Forms.GroupBox FootnotesBox;
        private System.Windows.Forms.ComboBox notesPositionSelector;
        private System.Windows.Forms.ComboBox notesLevelSelector;
        private System.Windows.Forms.Label notesLevelLabel;
        private System.Windows.Forms.Label notesPositionLabel;
        private System.Windows.Forms.ComboBox notesNumberingSelector;
        private System.Windows.Forms.Label notesNumberingLabel;
        private System.Windows.Forms.Label notesNumberPrefixLabel;
        private System.Windows.Forms.Label notesNumberSuffixLabel;
        private System.Windows.Forms.TextBox notesNumberSuffixValue;
        private System.Windows.Forms.TextBox notesNumberPrefixValue;
        private System.Windows.Forms.Label notesNumberingStartLabel;
        private System.Windows.Forms.MaskedTextBox notesNumberingStartValue;
        private System.Windows.Forms.GroupBox TextToSpeechBox;
        private System.Windows.Forms.Label TTSConfigFileLabel;
        private System.Windows.Forms.Button TTSConfigFileBrowse;
        private System.Windows.Forms.TextBox TTSConfigFilePath;
        private System.Windows.Forms.Label AzureRegionLabel;
        private System.Windows.Forms.ComboBox AzureRegionValue;
        private System.Windows.Forms.Label AzureKeyLabel;
        private System.Windows.Forms.TextBox AzureKeyValue;
    }
}