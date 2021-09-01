using System;
using System.IO;
using System.Xml;
using System.Net;
using System.Data;
using System.Drawing;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class About : Form
    {
        private string _currentVersion;

        /// <summary>
        /// Constructor which initializes the components
        /// Look and feel of UI
        /// </summary>
        public About()
        {
            InitializeComponent();
            _currentVersion = GetCurrentversion();
            this.versionNumLabel.Text = _currentVersion;
        }

        /// <summary>
        /// Function to check whether Update is there or not
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Updatesbutton_Click(object sender, EventArgs e)
        {
            try
            {
                string sourceURL = "http://openxml-daisy.sourceforge.net/DaisyTranslatorUpdates/version.xml";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sourceURL);
                request.Proxy.Credentials = CredentialCache.DefaultCredentials;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream receiveStream = response.GetResponseStream();
                StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);

                XmlDocument documnt = new XmlDocument();
                documnt.Load(readStream);
                String NewVersion = documnt.DocumentElement.FirstChild.InnerText;
                response.Close();
                readStream.Close();

                _currentVersion = GetCurrentversion();

                int Installedversion = Convert.ToInt16(_currentVersion.Replace(".", ""));
                int Availableversion = Convert.ToInt16(NewVersion.Replace(".", ""));

                string Download = sourceURL.Replace("version.xml", NewVersion + "/" + "DAISYTranslatorAddInForWordSetup-en.exe");
                if (Availableversion == Installedversion || Availableversion < Installedversion)
                {
                    MessageBox.Show("New version is not available", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (Availableversion > Installedversion)
                {
                    DialogResult dr = MessageBox.Show("New version of SaveAsDAISY Add-in is available.Do you want to download it now?", "SaveAsDAISY-New Version", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (dr == DialogResult.OK)
                        System.Diagnostics.Process.Start(Download);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"SaveAsDAISY",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            this.Close();
        }

        /// <summary>
        /// Function which gives the Current Version of the SaveAsDAISY
        /// </summary>
        /// <returns></returns>
        public string GetCurrentversion()
        {
            //string temp = Assembly.GetExecutingAssembly().Location;
            //Assembly obj = Assembly.LoadFile(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Converter.dll");
            //_currentVersion = obj.GetName().Version.ToString();
            _currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            return _currentVersion;
        }

        /// <summary>
        /// Function to Close the Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BttnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Function to close the Form on click of Esc button. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void About_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)27)
            {
                this.Close();
            }
        }
    }
}