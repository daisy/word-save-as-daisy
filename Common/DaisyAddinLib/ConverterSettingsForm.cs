using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Text;
using System.Xml.Xsl;
using System.Windows.Forms;
using System.Drawing;
using System.Resources;
using System.Xml.XPath;
using System.Reflection;
using System.Collections;
using System.IO.Packaging;
using System.ComponentModel;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class ConverterSettingsForm : Form
    {
        XmlDocument SettingsXml = new XmlDocument();
        String xmlfile_path;

        public ConverterSettingsForm()
        {
            InitializeComponent();
        }



        private void Daisysettingsfrm_Load(object sender, EventArgs e)
        {
            xmlfile_path = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SaveAsDAISY");

            if (!File.Exists(xmlfile_path + "\\DAISY_settingsVer21.xml"))
            {
            	ConverterSettings.CreateDefaultSettings(xmlfile_path);
            }

			SettingsXml.Load(xmlfile_path + "\\DAISY_settingsVer21.xml");
            XmlNode node = SettingsXml.SelectSingleNode("//Settings/PageNumbers[@value]");

            if (node.Attributes[0].Value == "Custom")
            {
                this.radiobtn_custom.Checked = true;
            }
            else
            {
                this.radiobtn_auto.Checked = true;
            }

            XmlNode charnode = SettingsXml.SelectSingleNode("//Settings/CharacterStyles[@value]");
            if (charnode.Attributes[0].Value == "True")
            {
                this.checkbox_translate.Checked = true;
            }
            else
            {
                this.checkbox_translate.Checked = false;
            }
            XmlNode imgnode = SettingsXml.LastChild["ImageSizes"];
            if (imgnode.Attributes[0].Value == "original")
            {

                this.radiobtn_originalimage.Checked = true;
                this.combobox_resample.Enabled = false;

            }
            else if (imgnode.Attributes[0].Value == "resize")
            {

                this.radiobtn_resize.Checked = true;
                this.combobox_resample.Enabled = false;
            }
            else if (imgnode.Attributes[0].Value == "resample")
            {
                this.radiobtn_resample.Checked = true;
                this.combobox_resample.Text = imgnode.Attributes[1].InnerXml;
                this.combobox_resample.Enabled = true;
            }
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            try
            {
				if (!File.Exists(xmlfile_path + "\\DAISY_settingsVer21.xml"))
				{
					ConverterSettings.CreateDefaultSettings(xmlfile_path);
				}

				SettingsXml.Load(xmlfile_path + "\\DAISY_settingsVer21.xml");

                XmlNode node = SettingsXml.SelectSingleNode("//Settings/PageNumbers[@value]");

                if (this.radiobtn_custom.Checked == true)
                {
                    node.Attributes[0].Value = "Custom";
                }
                else
                {
                    node.Attributes[0].Value = "Automatic";
                }
                XmlNode charnode = SettingsXml.SelectSingleNode("//Settings/CharacterStyles[@value]");
                if (this.checkbox_translate.Checked == true)
                {
                    charnode.Attributes[0].Value = "True";

                }
                else
                {
                    charnode.Attributes[0].Value = "False";
                }
                XmlNode imgnode = SettingsXml.LastChild["ImageSizes"];

                if (this.radiobtn_originalimage.Checked == true)
                {
                    imgnode.Attributes[0].Value = "original";

                }
                else if (this.radiobtn_resize.Checked == true)
                {
                    imgnode.Attributes[0].Value = "resize";

                }
                else if (this.radiobtn_resample.Checked == true)
                {

                    imgnode.Attributes[0].Value = "resample";
                    imgnode.Attributes[1].Value = this.combobox_resample.SelectedItem.ToString();

                }

                SettingsXml.Save(xmlfile_path + "\\DAISY_settingsVer21.xml");
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

       
    }
}