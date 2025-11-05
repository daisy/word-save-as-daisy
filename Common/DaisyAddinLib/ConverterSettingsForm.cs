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

        private static readonly Dictionary<string, string> AzureRegionDictionnary = new Dictionary<string, string>()
        {
            {"","" },
            {"East US", "eastus"},
            {"East US 2", "eastus2"},
            {"South Central US", "southcentralus"},
            {"West US 2", "westus2"},
            {"West US 3", "westus3"},
            {"Australia East", "australiaeast"},
            {"Southeast Asia", "southeastasia"},
            {"North Europe", "northeurope"},
            {"Sweden Central", "swedencentral"},
            {"UK South", "uksouth"},
            {"West Europe", "westeurope"},
            {"Central US", "centralus"},
            {"South Africa North", "southafricanorth"},
            {"Central India", "centralindia"},
            {"East Asia", "eastasia"},
            {"Japan East", "japaneast"},
            {"Korea Central", "koreacentral"},
            {"Canada Central", "canadacentral"},
            {"France Central", "francecentral"},
            {"Germany West Central", "germanywestcentral"},
            {"Norway East", "norwayeast"},
            {"Switzerland North", "switzerlandnorth"},
            {"UAE North", "uaenorth"},
            {"Brazil South", "brazilsouth"},
            {"Central US EUAP", "centraluseuap"},
            {"East US 2 EUAP", "eastus2euap"},
            {"Qatar Central", "qatarcentral"},
            {"Central US (Stage)", "centralusstage"},
            {"East US (Stage)", "eastusstage"},
            {"East US 2 (Stage)", "eastus2stage"},
            {"North Central US (Stage)", "northcentralusstage"},
            {"South Central US (Stage)", "southcentralusstage"},
            {"West US (Stage)", "westusstage"},
            {"West US 2 (Stage)", "westus2stage"},
            {"Asia", "asia"},
            {"Asia Pacific", "asiapacific"},
            {"Australia", "australia"},
            {"Brazil", "brazil"},
            {"Canada", "canada"},
            {"Europe", "europe"},
            {"France", "france"},
            {"Germany", "germany"},
            {"Global", "global"},
            {"India", "india"},
            {"Japan", "japan"},
            {"Korea", "korea"},
            {"Norway", "norway"},
            {"Singapore", "singapore"},
            {"South Africa", "southafrica"},
            {"Switzerland", "switzerland"},
            {"United Arab Emirates", "uae"},
            {"United Kingdom", "uk"},
            {"United States", "unitedstates"},
            {"United States EUAP", "unitedstateseuap"},
            {"East Asia (Stage)", "eastasiastage"},
            {"Southeast Asia (Stage)", "southeastasiastage"},
            {"Brazil US", "brazilus"},
            {"East US STG", "eastusstg"},
            {"North Central US", "northcentralus"},
            {"West US", "westus"},
            {"Jio India West", "jioindiawest"},
            {"devfabric", "devfabric"},
            {"West Central US", "westcentralus"},
            {"South Africa West", "southafricawest"},
            {"Australia Central", "australiacentral"},
            {"Australia Central 2", "australiacentral2"},
            {"Australia Southeast", "australiasoutheast"},
            {"Japan West", "japanwest"},
            {"Jio India Central", "jioindiacentral"},
            {"Korea South", "koreasouth"},
            {"South India", "southindia"},
            {"West India", "westindia"},
            {"Canada East", "canadaeast"},
            {"France South", "francesouth"},
            {"Germany North", "germanynorth"},
            {"Norway West", "norwaywest"},
            {"Switzerland West", "switzerlandwest"},
            {"UK West", "ukwest"},
            {"UAE Central", "uaecentral"},
            {"Brazil Southeast", "brazilsoutheast"},
        };

        private static readonly Dictionary<string, string> AzureRegionRevertDictionnary = new Dictionary<string, string>() {
            {"" , ""},
            {"eastus", "East US"},
            {"eastus2", "East US 2"},
            {"southcentralus", "South Central US"},
            {"westus2", "West US 2"},
            {"westus3", "West US 3"},
            {"australiaeast", "Australia East"},
            {"southeastasia", "Southeast Asia"},
            {"northeurope", "North Europe"},
            {"swedencentral", "Sweden Central"},
            {"uksouth", "UK South"},
            {"westeurope", "West Europe"},
            {"centralus", "Central US"},
            {"southafricanorth", "South Africa North"},
            {"centralindia", "Central India"},
            {"eastasia", "East Asia"},
            {"japaneast", "Japan East"},
            {"koreacentral", "Korea Central"},
            {"canadacentral", "Canada Central"},
            {"francecentral", "France Central"},
            {"germanywestcentral", "Germany West Central"},
            {"norwayeast", "Norway East"},
            {"switzerlandnorth", "Switzerland North"},
            {"uaenorth", "UAE North"},
            {"brazilsouth", "Brazil South"},
            {"centraluseuap", "Central US EUAP"},
            {"eastus2euap", "East US 2 EUAP"},
            {"qatarcentral", "Qatar Central"},
            {"centralusstage", "Central US (Stage)"},
            {"eastusstage", "East US (Stage)"},
            {"eastus2stage", "East US 2 (Stage)"},
            {"northcentralusstage", "North Central US (Stage)"},
            {"southcentralusstage", "South Central US (Stage)"},
            {"westusstage", "West US (Stage)"},
            {"westus2stage", "West US 2 (Stage)"},
            {"asia", "Asia"},
            {"asiapacific", "Asia Pacific"},
            {"australia", "Australia"},
            {"brazil", "Brazil"},
            {"canada", "Canada"},
            {"europe", "Europe"},
            {"france", "France"},
            {"germany", "Germany"},
            {"global", "Global"},
            {"india", "India"},
            {"japan", "Japan"},
            {"korea", "Korea"},
            {"norway", "Norway"},
            {"singapore", "Singapore"},
            {"southafrica", "South Africa"},
            {"switzerland", "Switzerland"},
            {"uae", "United Arab Emirates"},
            {"uk", "United Kingdom"},
            {"unitedstates", "United States"},
            {"unitedstateseuap", "United States EUAP"},
            {"eastasiastage", "East Asia (Stage)"},
            {"southeastasiastage", "Southeast Asia (Stage)"},
            {"brazilus", "Brazil US"},
            {"eastusstg", "East US STG"},
            {"northcentralus", "North Central US"},
            {"westus", "West US"},
            {"jioindiawest", "Jio India West"},
            {"devfabric", "devfabric"},
            {"westcentralus", "West Central US"},
            {"southafricawest", "South Africa West"},
            {"australiacentral", "Australia Central"},
            {"australiacentral2", "Australia Central 2"},
            {"australiasoutheast", "Australia Southeast"},
            {"japanwest", "Japan West"},
            {"jioindiacentral", "Jio India Central"},
            {"koreasouth", "Korea South"},
            {"southindia", "South India"},
            {"westindia", "West India"},
            {"canadaeast", "Canada East"},
            {"francesouth", "France South"},
            {"germanynorth", "Germany North"},
            {"norwaywest", "Norway West"},
            {"switzerlandwest", "Switzerland West"},
            {"ukwest", "UK West"},
            {"uaecentral", "UAE Central"},
            {"brazilsoutheast", "Brazil Southeast"},
        };

        private void Daisysettingsfrm_Load(object sender, EventArgs e)
        {
            notesNumberingStartValue.Mask = "000";

            if (GlobaleSettings.PagenumStyle == ConverterSettings.PageNumberingChoice.Enum.Custom)
            {
                this.radiobtn_custom.Checked = true;
            }
            else
            {
                this.radiobtn_auto.Checked = true;
            }

            if (GlobaleSettings.CharacterStyle)
            {
                this.checkbox_translate.Checked = true;
            }
            else
            {
                this.checkbox_translate.Checked = false;
            }

            if (GlobaleSettings.ImageOption == ConverterSettings.ImageOptionChoice.Enum.Original)
            {

                this.radiobtn_originalimage.Checked = true;
                this.combobox_resample.Enabled = false;

            }
            else if (GlobaleSettings.ImageOption == ConverterSettings.ImageOptionChoice.Enum.Resize)
            {

                this.radiobtn_resize.Checked = true;
                this.combobox_resample.Enabled = false;
            }
            else if (GlobaleSettings.ImageOption == ConverterSettings.ImageOptionChoice.Enum.Resample)
            {
                this.radiobtn_resample.Checked = true;
                this.combobox_resample.Text = GlobaleSettings.ImageResamplingValue.ToString();
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
                    this.notesNumberingSelector.SelectedItem = numberingPair.Key;
                    // Disable note numbering start value field if numbering scheme is not set to "Number"
                    this.notesNumberingStartValue.Enabled = numberingPair.Value == ConverterSettings.FootnotesNumberingChoice.Enum.Number;
                    // Disable level selection for end of pages value
                    //footnotesLevelSelector.Enabled = (numberingPair.Value != ConverterSettings.FootnotesPositionChoice.Enum.Page);
                }
            }
            notesNumberingStartValue.Text = GlobaleSettings.FootnotesStartValue.ToString();
            notesNumberPrefixValue.Text = GlobaleSettings.FootnotesNumberingPrefix;
            notesNumberSuffixValue.Text = GlobaleSettings.FootnotesNumberingSuffix;
            //TTSConfigFilePath.Text = GlobaleSettings.TTSConfigFile;
            //AzureKeyValue.Text = GlobaleSettings.AzureSpeechKey;
            //AzureRegionValue.SelectedItem = AzureRegionRevertDictionnary[GlobaleSettings.AzureSpeechRegion ?? ""];

            this.notesNumberingStartValue.Enabled = notesNumberingMap[(string)notesNumberingSelector.SelectedItem] == ConverterSettings.FootnotesNumberingChoice.Enum.Number;

            this.DontNotify.Checked = GlobaleSettings.DontNotifySponsorship;

        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            try
            {
                // Update fields
                GlobaleSettings.PagenumStyle = this.radiobtn_custom.Checked 
                    ? ConverterSettings.PageNumberingChoice.Enum.Custom 
                    : ConverterSettings.PageNumberingChoice.Enum.Automatic;
                GlobaleSettings.CharacterStyle = this.checkbox_translate.Checked;

                if (this.radiobtn_originalimage.Checked == true)
                {
                    GlobaleSettings.ImageOption = ConverterSettings.ImageOptionChoice.Enum.Original;

                }
                else if (this.radiobtn_resize.Checked == true)
                {
                    GlobaleSettings.ImageOption = ConverterSettings.ImageOptionChoice.Enum.Resize;

                }
                else if (this.radiobtn_resample.Checked == true)
                {

                    GlobaleSettings.ImageOption = ConverterSettings.ImageOptionChoice.Enum.Resample;
                    GlobaleSettings.ImageResamplingValue = (ConverterSettings.ImageResamplingChoice.Enum)this.combobox_resample.SelectedItem;

                }


                string selectedPosition = (string)notesPositionSelector.SelectedItem;
                ConverterSettings.FootnotesPositionChoice.Enum positionValue;
                if (!positionsMap.TryGetValue(selectedPosition, out positionValue))
                {
                    positionValue = ConverterSettings.FootnotesPositionChoice.Enum.Page;
                }
                GlobaleSettings.FootnotesPosition = positionValue;


                string selectedLevel = (string)notesLevelSelector.SelectedItem;
                ConverterSettings.FootnotesLevelChoice.Enum levelValue;
                if (!levelsMap.TryGetValue(selectedLevel, out levelValue))
                {
                    levelValue = ConverterSettings.FootnotesLevelChoice.Enum.Inlined;
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
                
                //GlobaleSettings.TTSConfigFile = TTSConfigFilePath.Text;
                //GlobaleSettings.AzureSpeechKey = AzureKeyValue.Text;
                //GlobaleSettings.AzureSpeechRegion = AzureRegionDictionnary[(string)(AzureRegionValue.SelectedItem ?? "")];

                // Save
                GlobaleSettings.Save();
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

        private void TTSConfigFileBrowse_Click(object sender, EventArgs e)
        {
            if (TTSConfigFileSelect.ShowDialog(this) == DialogResult.OK) {
                if(TTSConfigFileSelect.FileName.Length > 0) {
                    GlobaleSettings.TTSConfigFile = TTSConfigFilePath.Text = TTSConfigFileSelect.FileName;
                }
            };
        }

        private void DontNotify_CheckedChanged(object sender, EventArgs e)
        {
            GlobaleSettings.DontNotifySponsorship = DontNotify.Checked;
        }
    }
}