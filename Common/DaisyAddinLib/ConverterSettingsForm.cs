using System;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class ConverterSettingsForm : Form
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;


        public ConverterSettingsForm()
        {
            InitializeComponent();
        }



        private void Daisysettingsfrm_Load(object sender, EventArgs e)
        {
            notesNumberingStartValue.Mask = "000";
            
            if (GlobaleSettings.PagenumStyle == "Custom")
            {
                this.radiobtn_custom.Checked = true;
            }
            else
            {
                this.radiobtn_auto.Checked = true;
            }

            if (GlobaleSettings.CharacterStyle == "True")
            {
                this.checkbox_translate.Checked = true;
            }
            else
            {
                this.checkbox_translate.Checked = false;
            }

            if (GlobaleSettings.ImageOption == "original")
            {

                this.radiobtn_originalimage.Checked = true;
                this.combobox_resample.Enabled = false;

            }
            else if (GlobaleSettings.ImageOption == "resize")
            {

                this.radiobtn_resize.Checked = true;
                this.combobox_resample.Enabled = false;
            }
            else if (GlobaleSettings.ImageOption == "resample")
            {
                this.radiobtn_resample.Checked = true;
                this.combobox_resample.Text = GlobaleSettings.ImageResamplingValue;
                this.combobox_resample.Enabled = true;
            }


            // NP 20220428 : Footnotes settings
            // Found back the key associated to the value
            foreach (var levelPair in levelsMap)
            {
                if (levelPair.Value == GlobaleSettings.FootnotesLevel)
                {
                    this.notesLevelSelector.SelectedItem = levelPair.Key;
                }
            }
            // Found back the key associated to the value
            foreach (var positionsPair in positionsMap)
            {
                if (positionsPair.Value == GlobaleSettings.FootnotesPosition)
                {
                    this.notesPositionSelector.SelectedItem = positionsPair.Key;
                    // Disable level selection for end of pages value
                    notesLevelSelector.Enabled = (positionsPair.Value != ConverterSettings.FootnotesPositionChoice.Enum.Page);
                }
            }
            // Found back the key associated to the value
            foreach (var numberingPair in notesNumberingMap)
            {
                if (numberingPair.Value == GlobaleSettings.FootnotesNumbering)
                {
                    this.notesPositionSelector.SelectedItem = numberingPair.Key;
                    this.notesNumberingStartValue.Enabled = numberingPair.Value == ConverterSettings.FootnotesNumberingChoice.Enum.Number;
                    // Disable level selection for end of pages value
                    //footnotesLevelSelector.Enabled = (numberingPair.Value != ConverterSettings.FootnotesPositionChoice.Enum.Page);
                }
            }
            notesNumberingStartValue.Text = GlobaleSettings.FootnotesStartValue.ToString();
            notesNumberPrefixValue.Text = GlobaleSettings.FootnotesNumberingPrefix;
            notesNumberSuffixValue.Text = GlobaleSettings.FootnotesNumberingSuffix;
            
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            try {
                // Update fields
                GlobaleSettings.PagenumStyle = this.radiobtn_custom.Checked ? "Custom" : "Automatic";
                GlobaleSettings.CharacterStyle = this.checkbox_translate.Checked ? "True" : "False";
                GlobaleSettings.CharacterStyle = this.checkbox_translate.Checked ? "True" : "False";

                if (this.radiobtn_originalimage.Checked == true)
                {
                    GlobaleSettings.ImageOption = "original";

                }
                else if (this.radiobtn_resize.Checked == true)
                {
                    GlobaleSettings.ImageOption = "resize";

                }
                else if (this.radiobtn_resample.Checked == true)
                {

                    GlobaleSettings.ImageOption = "resample";
                    GlobaleSettings.ImageResamplingValue = this.combobox_resample.SelectedItem.ToString();

                }


                string selectedPosition = (string)notesPositionSelector.SelectedItem;
                ConverterSettings.FootnotesPositionChoice.Enum positionValue;
                if (!positionsMap.TryGetValue(selectedPosition, out positionValue))
                {
                    positionValue = ConverterSettings.FootnotesPositionChoice.Enum.Page;
                }
                GlobaleSettings.FootnotesPosition = positionValue;


                string selectedLevel = (string)notesLevelSelector.SelectedItem;
                int levelValue;
                if (!levelsMap.TryGetValue(selectedLevel, out levelValue))
                {
                    levelValue = 0;
                }
                GlobaleSettings.FootnotesLevel = levelValue;

                string selectedNumbering = (string)notesNumberingSelector.SelectedItem;
                ConverterSettings.FootnotesNumberingChoice.Enum numberingValue;
                if (!notesNumberingMap.TryGetValue(selectedNumbering, out numberingValue))
                {
                    numberingValue = ConverterSettings.FootnotesNumberingChoice.Enum.None;
                }
                GlobaleSettings.FootnotesNumbering = numberingValue;
                GlobaleSettings.FootnotesStartValue = int.Parse(notesNumberingStartValue.Text);
                GlobaleSettings.FootnotesNumberingPrefix = notesNumberPrefixValue.Text;
                GlobaleSettings.FootnotesNumberingSuffix = notesNumberSuffixValue.Text;


                // Save
                GlobaleSettings.save();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Submit Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void radiobtn_resample_MouseClick(object sender, MouseEventArgs e)
        {
            if (this.radiobtn_resample.Checked == true)
            {
                this.combobox_resample.Enabled = true;

            }
        }

        private void radiobtn_resample_Click(object sender, EventArgs e)
        {
            if (this.radiobtn_resample.Checked == true)
            {
                this.combobox_resample.Text = this.combobox_resample.Items[1].ToString();
                this.combobox_resample.Enabled = true;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radiobtn_originalimage_Click(object sender, EventArgs e)
        {
            if (this.radiobtn_originalimage.Checked == true)
            {
                this.combobox_resample.Enabled = false;
            }
        }

        private void radiobtn_resize_Click(object sender, EventArgs e)
        {
            if (this.radiobtn_resize.Checked == true)
            {
                this.combobox_resample.Enabled = false;
            }
        }

        

        private void footnotesLevelSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*string selectedItem = (string)footNotesLevel.SelectedItem;

            string value;
            if(!level.TryGetValue(selectedItem, out value))
            {
                value = "0";
            }*/

        }

        private void footnotesPositionSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            notesLevelSelector.Enabled = ((string)notesPositionSelector.SelectedItem != "End of pages");
            
           /* string selectedItem = (string)footNotesLevel.SelectedItem;

            string value;
            if (!level.TryGetValue(selectedItem, out value))
            {
                value = "page";
            }*/
        }

        private void notesNumberingSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.notesNumberingStartValue.Enabled = notesNumberingMap[(string)notesNumberingSelector.SelectedItem] == ConverterSettings.FootnotesNumberingChoice.Enum.Number;
        }

    }
}