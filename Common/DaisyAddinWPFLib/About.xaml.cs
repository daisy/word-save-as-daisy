using System;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour About.xaml
    /// </summary>
    public partial class About : Window
    {

        public About()
        {
            try {
                InitializeComponent();
            } catch (Exception ex) {
                MessageBox.Show(
                    "Error initializing About dialog: " + ex.Message,
                    "SaveAsDAISY",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                throw ex;
            }
            _currentVersion = GetCurrentversion();
            AddinVersion.Text = string.Format(this.AddinVersion.Text.ToString(), _currentVersion);
        }

        private string _currentVersion;


        private int computeVersionComparator(string version)
        {
            Regex versionSeek = new Regex(@"(\d+)\.(\d+)\.(\d+)");
            Match found = versionSeek.Match(version);
            if (found.Success) {
                int major = Int32.Parse(found.Groups[1].Value);
                int minor = Int32.Parse(found.Groups[2].Value);
                int patch = Int32.Parse(found.Groups[3].Value);
                return patch + ((minor + 1) * 1000) + ((major + 1) * 1000000);
            }
            return 0;
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
        /// Function to check whether Update is there or not
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Updatesbutton_Click(object sender, RoutedEventArgs e)
        {
            try {
                string sourceURL = "https://github.com/daisy/word-save-as-daisy/releases/latest";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sourceURL);
                request.AllowAutoRedirect = true;
                request.Proxy.Credentials = CredentialCache.DefaultCredentials;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string lastTag = response.ResponseUri.ToString().Split('/').Last();
                response.Close();

                if (computeVersionComparator(lastTag) > computeVersionComparator(GetCurrentversion())) {
                    MessageBoxResult result = MessageBox.Show(
                        "New version of SaveAsDAISY Add-in is available. Do you want to download it now?\r\n" +
                        "(You will be redirected to the latest release download page)",
                        "SaveAsDAISY-New Version",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Information
                    );
                    if (result == MessageBoxResult.OK)
                        System.Diagnostics.Process.Start("https://github.com/daisy/word-save-as-daisy/releases/latest");
                } else {
                    MessageBox.Show(
                        "You already have the latest version of the plugin.",
                        "SaveAsDAISY",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                }
            }
            catch (Exception ex) {
                MessageBox.Show(
                    ex.Message,
                    "SaveAsDAISY",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }


        /// <summary>
        /// Function to Close the Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        /// <summary>
        /// Function to close the Form on click of Esc button. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void About_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (e.KeyChar == (char)27) {
        //        this.Close();
        //    }
        //}
    }
}
