using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Xml;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour Settings.xaml
    /// </summary>
    public partial class SettingsForm : Window
    {
        private static ConverterSettings GlobaleSettings = ConverterSettings.Instance;

        private static readonly Dictionary<string, PageNumberingChoice.Enum> PageNumberingChoices = new Dictionary<string, PageNumberingChoice.Enum>
        {
            { "Compute from 'Page number (DAISY)' elements", PageNumberingChoice.Enum.Custom },
            { "Compute from Word rendered page breaks", PageNumberingChoice.Enum.Automatic },
        };

        private static readonly Dictionary<string, ImageResamplingChoice.Enum> ResamplingValues = new Dictionary<string, ImageResamplingChoice.Enum>
        {
            { "72 dpi",  ImageResamplingChoice.Enum.dpi_72 },
            { "96 dpi",  ImageResamplingChoice.Enum.dpi_96 },
            { "120 dpi", ImageResamplingChoice.Enum.dpi_120 },
            { "150 dpi", ImageResamplingChoice.Enum.dpi_150 },
            { "300 dpi", ImageResamplingChoice.Enum.dpi_300 },
        };

        private static readonly Dictionary<string, FootnotesLevelChoice.Enum> NotesLevelChoices = new Dictionary<string, FootnotesLevelChoice.Enum>
        {
            { "Note reference level", FootnotesLevelChoice.Enum.Inlined},
            { "First or nearest level",  FootnotesLevelChoice.Enum.Level_1},
            { "Second or nearest level", FootnotesLevelChoice.Enum.Level_2},
            { "Third or nearest level", FootnotesLevelChoice.Enum.Level_3},
            { "Fourth or nearest level", FootnotesLevelChoice.Enum.Level_4},
            { "Fifth or nearest level", FootnotesLevelChoice.Enum.Level_5},
            { "Sixth or nearest level", FootnotesLevelChoice.Enum.Level_6}
        };

        

        private static readonly Dictionary<string, FootnotesNumberingChoice.Enum> NotesNumberingChoices = new Dictionary<string, FootnotesNumberingChoice.Enum>
        {
            {  "Don't add a numeric prefix", FootnotesNumberingChoice.Enum.None },
            {  "Recompute a numeric prefix", FootnotesNumberingChoice.Enum.Number }
        };

        private static readonly Dictionary<string, ImageOptionChoice.Enum> ImageSizeChoices = new Dictionary<string, ImageOptionChoice.Enum>
        {
            {  "Keep original images", ImageOptionChoice.Enum.Original },
            {  "Resize the images", ImageOptionChoice.Enum.Resize },
            {  "Resample the images", ImageOptionChoice.Enum.Resample }
        };

        private static readonly Dictionary<string, FootnotesPositionChoice.Enum> NotesPositionChoices = new Dictionary<string, FootnotesPositionChoice.Enum>
        {
            {"End of pages",FootnotesPositionChoice.Enum.Page },
            {"Inlined in levels", FootnotesPositionChoice.Enum.Inline},
            {"End of levels", FootnotesPositionChoice.Enum.End}
        };

        

        public SettingsForm()
        {
            InitializeComponent();

            TranslateCharacteStyles.IsChecked = GlobaleSettings.CharacterStyle;
            DisableSponsorhip.IsChecked = GlobaleSettings.DontNotifySponsorship;

            PageNumbering.ItemsSource = PageNumberingChoices.Keys;
            PageNumbering.SelectedIndex = PageNumberingChoices.Values.ToList().IndexOf(GlobaleSettings.PagenumStyle);

            ImageSizeOptions.ItemsSource = ImageSizeChoices.Keys;
            ImageSizeOptions.SelectedIndex= ImageSizeChoices.Values.ToList().IndexOf(GlobaleSettings.ImageOption);

            Resampling.IsEnabled = GlobaleSettings.ImageOption == ImageOptionChoice.Enum.Resample;
            Resampling.ItemsSource = ResamplingValues.Keys;
            Resampling.SelectedIndex = ResamplingValues.Values.ToList().IndexOf(GlobaleSettings.ImageResamplingValue);

            NotesPosition.ItemsSource = NotesPositionChoices.Keys;
            NotesPosition.SelectedIndex = NotesPositionChoices.Values.ToList().IndexOf(GlobaleSettings.FootnotesPosition);

            NotesNumbering.ItemsSource = NotesNumberingChoices.Keys;
            NotesNumbering.SelectedIndex = NotesNumberingChoices.Values.ToList().IndexOf(GlobaleSettings.FootnotesNumbering);

            NotesLevel.ItemsSource = NotesLevelChoices.Keys;
            NotesLevel.SelectedIndex = NotesLevelChoices.Values.ToList().IndexOf(GlobaleSettings.FootnotesLevel);

            StartNumber.IsEnabled = GlobaleSettings.FootnotesNumbering != FootnotesNumberingChoice.Enum.None;
            StartNumber.Text = GlobaleSettings.FootnotesStartValue.ToString().PadLeft(3,'_');
            NumberPrefix.Text = GlobaleSettings.FootnotesNumberingPrefix;
            NotesTextPrefix.Text = GlobaleSettings.FootnotesNumberingSuffix;

            OpenOfficeTemplate.Text = GlobaleSettings.OTTTemplateFile;

            UseWebserviceRunner.IsChecked = GlobaleSettings.UseWebserviceRunner;

            UseDAISYPipelineApp.IsEnabled = UseWebserviceRunner.IsChecked == true;
            UseDAISYPipelineApp.IsChecked = UseWebserviceRunner.IsChecked == true && ConverterHelper.PipelineAppIsInstalled() && GlobaleSettings.UseDAISYPipelineApp;
            PreferredVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
            TTSEngines.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
            BrowseVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;

            TTSConfigFile.Text = GlobaleSettings.TTSConfigFile;
            TTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked == false;
            BrowseTTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked == false;
            OpenPipelineProperties.IsEnabled = UseDAISYPipelineApp.IsChecked == false;
        }

        private void ImageSizeOptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            Resampling.IsEnabled = ImageSizeChoices.Values.ToList()[ImageSizeOptions.SelectedIndex] == ImageOptionChoice.Enum.Resample;
        }

        private static readonly Regex _regex = new Regex("_*[^0-9]+"); //regex that matches disallowed text
        private static bool IsTextAllowed(string text)
        {
            return !_regex.IsMatch(text);
        }

        private void StartNumber_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = IsTextAllowed(e.Text);
        }

        private void StartNumber_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string))) {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!IsTextAllowed(text)) {
                    e.CancelCommand();
                }
            } else {
                e.CancelCommand();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            this.Close();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (UseDAISYPipelineApp.IsChecked == false && !string.IsNullOrWhiteSpace(TTSConfigFile.Text)) {
                try {
                    if(!System.IO.File.Exists(TTSConfigFile.Text.Trim())) {
                        throw new Exception("The provided TTS config file does not exist.");
                    }
                    XmlDocument config = new XmlDocument();
                    config.Load(TTSConfigFile.Text);
                    XmlNode root = config.DocumentElement;
                    if (root == null || root.Name != "config") {
                        throw new Exception("The root element should be named 'config'.");
                    }
                    // If any child node under config is not a voice or a lexicon element, throw an exception
                    foreach (XmlNode child in root.ChildNodes) {
                        if (child.Name != "voice" && child.Name != "lexicon") {
                            throw new Exception("Only 'voice' or 'lexicon' elements are allowed under 'config' (found " + child.Name + ")");
                        }
                    }
                }
                catch (Exception ex) {
                    MessageBoxResult res =  MessageBox.Show(
                        "The provided TTS config file is not valid: " + ex.Message + 
                        "\r\n" +
                        "Do you want to see the TTS configuration file documentation and fix it ?\r\n" +
                        "- Click Yes to open the documentation\r\n" +
                        "- Click No to reset your config file selection\r\n" +
                        "- Click Cancel to continue editing your settings", "Invalid TTS Config File", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
                    switch (res) {
                        case MessageBoxResult.Yes:
                            Process.Start(new ProcessStartInfo("https://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/") { UseShellExecute = true });
                            return;
                        case MessageBoxResult.No:
                            TTSConfigFile.Text = "";
                            return; 
                        case MessageBoxResult.Cancel:
                            return;
                    }
                }
                
            }
            try {
                // Update fields
                GlobaleSettings.PagenumStyle = PageNumberingChoices.Values.ToList()[PageNumbering.SelectedIndex];
                GlobaleSettings.CharacterStyle = TranslateCharacteStyles.IsChecked == true;
                GlobaleSettings.DontNotifySponsorship = DisableSponsorhip.IsChecked == true;
                GlobaleSettings.ImageOption = ImageSizeChoices.Values.ToList()[ImageSizeOptions.SelectedIndex];
                GlobaleSettings.ImageResamplingValue = Resampling.SelectedIndex >= 0 ? ResamplingValues.Values.ToList()[Resampling.SelectedIndex] : ImageResamplingChoice.Enum.dpi_72; // Default to 72 DPI if not selected
                GlobaleSettings.FootnotesPosition = NotesPositionChoices.Values.ToList()[NotesPosition.SelectedIndex];
                GlobaleSettings.FootnotesLevel = NotesLevelChoices.Values.ToList()[NotesLevel.SelectedIndex];
                GlobaleSettings.FootnotesNumbering = NotesNumberingChoices.Values.ToList()[NotesNumbering.SelectedIndex];
                GlobaleSettings.FootnotesStartValue = int.Parse(StartNumber.Text.Trim().Replace("_",""));
                GlobaleSettings.FootnotesNumberingPrefix = NumberPrefix.Text;
                GlobaleSettings.FootnotesNumberingSuffix = NotesTextPrefix.Text;
                GlobaleSettings.UseDAISYPipelineApp = UseDAISYPipelineApp.IsChecked == true;
                GlobaleSettings.TTSConfigFile = TTSConfigFile.Text.Trim();
                GlobaleSettings.OTTTemplateFile = OpenOfficeTemplate.Text.Trim();
                GlobaleSettings.UseWebserviceRunner = UseWebserviceRunner.IsChecked == true;
                // Save
                GlobaleSettings.Save();
                DialogResult = true;
                this.Close();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not save settings", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void NotesNumbering_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            StartNumber.IsEnabled = NotesNumberingChoices.Values.ToList()[NotesNumbering.SelectedIndex] != FootnotesNumberingChoice.Enum.None;
        }

        
        //private void UseWebserviceInteractions_Checked(object sender, RoutedEventArgs e)
        //{

        //}
        private void UseDAISYPipelineApp_Checked(object sender, RoutedEventArgs e)
        {
            if(!ConverterHelper.PipelineAppIsInstalled()) {
                var result = MessageBox.Show(
                    "The DAISY Pipeline application is not installed on this system. " +
                    "Do you want to install it ?",
                    "DAISY Pipeline App Not Found",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning
                );
                if(result == MessageBoxResult.Yes) {
                    Process.Start(ConverterHelper.PipelineAppReleaseURL);
                }
                UseDAISYPipelineApp.IsChecked = false;
                PreferredVoices.IsEnabled = false;
                TTSEngines.IsEnabled = false;
                BrowseVoices.IsEnabled = false;
                TTSConfigFile.IsEnabled = true;
                BrowseTTSConfigFile.IsEnabled = true;
                OpenPipelineProperties.IsEnabled = true;
            } else {
                GlobaleSettings.UseDAISYPipelineApp = true;
                PreferredVoices.IsEnabled = true;
                TTSEngines.IsEnabled = true;
                BrowseVoices.IsEnabled = true;

                TTSConfigFile.IsEnabled = false;
                BrowseTTSConfigFile.IsEnabled = false;
                OpenPipelineProperties.IsEnabled = false;
            }
        }
        private void UseDAISYPipelineApp_Unchecked(object sender, RoutedEventArgs e)
        {
            PreferredVoices.IsEnabled = false;
            TTSEngines.IsEnabled = false;
            BrowseVoices.IsEnabled = false;

            TTSConfigFile.IsEnabled = true;
            BrowseTTSConfigFile.IsEnabled = true;

        }

        private void PreferredVoices_Click(object sender, RoutedEventArgs e)
        {
            try {
                WebserviceRunner.StartDAISYPipelineAppWebservice();
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.PipelineAppPath,
                    Arguments = "preferred-voices",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(startInfo);
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not open preferred voices : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void TTSEngines_Click(object sender, RoutedEventArgs e)
        {
            try {
                WebserviceRunner.StartDAISYPipelineAppWebservice();
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.PipelineAppPath,
                    Arguments = "external-services",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(startInfo);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not open TTS engines settings : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BrowseVoices_Click(object sender, RoutedEventArgs e)
        {

            try {
                WebserviceRunner.StartDAISYPipelineAppWebservice();
                var startInfo = new ProcessStartInfo
                {
                    FileName = ConverterHelper.PipelineAppPath,
                    Arguments = "browse-voices",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(startInfo);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not browse voices : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }



        private void BrowseTTSConfigFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
           
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "XML file (*.xml)|*.xml";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true) {

                try {
                    
                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();
                    string fileContent = "";
                    using (StreamReader reader = new StreamReader(fileStream)) {
                        fileContent = reader.ReadToEnd();
                    }
                    XmlDocument config = new XmlDocument();
                    config.LoadXml(fileContent);
                    XmlNode root = config.DocumentElement;
                    if (root == null || root.Name != "config") {
                        throw new Exception("The root element should be named 'config'.");
                    }
                    // If any child node under config is not a voice or a lexicon element, throw an exception
                    foreach (XmlNode child in root.ChildNodes) {
                        if (child.Name != "voice" && child.Name != "lexicon") {
                            throw new Exception("Only 'voice' or 'lexicon' elements are allowed under 'config' (found " + child.Name + ")");
                        }
                    }
                    TTSConfigFile.Text = openFileDialog.FileName;
                }
                catch (Exception ex) {
                    MessageBoxResult res = MessageBox.Show(
                        "The provided TTS config file is not valid: " + ex.Message +
                        "\r\n" +
                        "Do you want to see the TTS configuration file documentation ?", "Invalid TTS Config File", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                    switch (res) {
                        case MessageBoxResult.Yes:
                            Process.Start(new ProcessStartInfo("https://daisy.github.io/pipeline/Get-Help/User-Guide/Text-To-Speech/") { UseShellExecute = true });
                            return;
                        default:
                            return;
                    }
                }

            }
            
        }

        private void BrowseOpenOfficeTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "OTT file (*.ott)|*.ott";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {

                OpenOfficeTemplate.Text = openFileDialog.FileName;
                
            }

        }

        private void BrowseResultsFolder_Click(object sender, RoutedEventArgs e)
        {

        }

        private void UseWebserviceRunner_Checked(object sender, RoutedEventArgs e)
        {
            // If it is not already enabled, confirm with the user that they want to enable it, as it may require allowing firewall access which can cause 'Access Denied' errors if not allowed
            if (!GlobaleSettings.UseWebserviceRunner)
            {
                var result = MessageBox.Show(
                      "The use of webservice interactions can speed up consecutive conversions of a document, but requires to have access to firewall security notices to allow the engine webservice.\r\n" +
                      "If you don't have this acces, you may encounter an 'Access Denied' error when trying to convert a document.\r\n" +
                      "Do you still want to enable the webservice interactions ?",
                      "Activating Webservice Interactions",
                      MessageBoxButton.YesNo,
                      MessageBoxImage.Warning
                );
                if (result == MessageBoxResult.Yes)
                {   
                    UseWebserviceRunner.IsChecked = true;
                    UseDAISYPipelineApp.IsEnabled = true;
                    PreferredVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                    TTSEngines.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                    BrowseVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                    TTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked != true;
                    BrowseTTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked != true;

                }
                else
                {
                    UseWebserviceRunner.IsChecked = false;
                    UseDAISYPipelineApp.IsEnabled = false;
                    PreferredVoices.IsEnabled = false;
                    TTSEngines.IsEnabled = false;
                    BrowseVoices.IsEnabled = false;
                    TTSConfigFile.IsEnabled = true;
                    BrowseTTSConfigFile.IsEnabled = true;
                }
            } else
            {
                UseWebserviceRunner.IsChecked = true;
                UseDAISYPipelineApp.IsEnabled = true;
                PreferredVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                TTSEngines.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                BrowseVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
                TTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked != true;
                BrowseTTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked != true;
            }
        }

        private void UseWebserviceRunner_Unchecked(object sender, RoutedEventArgs e)
        {
            UseWebserviceRunner.IsChecked = false;
            UseDAISYPipelineApp.IsEnabled = false;
            PreferredVoices.IsEnabled = false;
            TTSEngines.IsEnabled = false;
            BrowseVoices.IsEnabled = false;
            TTSConfigFile.IsEnabled = true;
            BrowseTTSConfigFile.IsEnabled = true;
        }

        private void OpenPipelineProperties_Click(object sender, RoutedEventArgs e)
        {
            // get current selected runtime descriptors
            if(UseWebserviceRunner.IsChecked != true)
            {
                try
                {
                    var dialog = new Pipeline2Properties(JNIWrapperRunner.GetInstance().GetSettableProperties(), PipelineUserProperties.Instance.Items);
                    if (dialog.ShowDialog() == true)
                    {
                        PipelineUserProperties.Instance.ReplaceBy(dialog.UpdatedProperties);
                        PipelineUserProperties.Instance.Save();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Could not open pipeline properties : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            } else {
                var dialog = new Pipeline2Properties(WebserviceRunner.GetInstance().GetSettableProperties(), PipelineUserProperties.Instance.Items);
                if (dialog.ShowDialog() == true)
                {
                    PipelineUserProperties.Instance.UpdateOrAddRange(dialog.UpdatedProperties);
                    PipelineUserProperties.Instance.Save();
                }
            }
        }
    }
}
