using System;
using System.IO;
using System.Xml;
using System.Net;
using System.Drawing;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;

namespace Daisy.SaveAsDAISY.Conversion
{
    public partial class About : Form
    {
        private string _currentVersion;

        private int computeVersionComparator(string version)
        {
            Regex versionSeek = new Regex(@"(\d+)\.(\d+)\.(\d+)");
            Match found = versionSeek.Match(version);
            if (found.Success)
            {
                int major = Int32.Parse(found.Groups[1].Value);
                int minor = Int32.Parse(found.Groups[2].Value);
                int patch = Int32.Parse(found.Groups[3].Value);
                return patch + ((minor + 1) * 1000) + ((major + 1) * 1000000);
            }
            return 0;
        }

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

                string sourceURL = "https://github.com/daisy/word-save-as-daisy/releases/latest";
                // from https://stackoverflow.com/questions/10822509/the-request-was-aborted-could-not-create-ssl-tls-secure-channel
                
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;
                // allows for validation of SSL conversations
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sourceURL);
                request.AllowAutoRedirect = true;
                request.Proxy.Credentials = CredentialCache.DefaultCredentials;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string lastTag = response.ResponseUri.ToString().Split('/').Last();
                response.Close();

                if (computeVersionComparator(lastTag) > computeVersionComparator(GetCurrentversion()))
                {
                    DialogResult dr = MessageBox.Show(
                        "New version of SaveAsDAISY Add-in is available.Do you want to download it now?\r\n" +
                            "(You will be redirected to the latest release download page)",
                        "SaveAsDAISY-New Version",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Information
                    );
                    if (dr == DialogResult.OK)
                        System.Diagnostics.Process.Start("https://github.com/daisy/word-save-as-daisy/releases/latest");

                }
                else
                {
                    MessageBox.Show("Your already have the latest version of the plugin.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"SaveAsDAISY",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            
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