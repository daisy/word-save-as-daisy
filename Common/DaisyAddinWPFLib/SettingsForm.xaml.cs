using Daisy.SaveAsDAISY.Conversion;
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

        private static readonly List<int> ResamplingValues = new List<int>
        {
            72, 96, 120, 150, 300
        };

        private static readonly Dictionary<string, int> NotesLevelChoices = new Dictionary<string, int>
        {
            { "Note reference level", 0},
            { "First or nearest level", 1},
            { "Second or nearest level", 2},
            { "Third or nearest level", 3},
            { "Fourth or nearest level", 4},
            { "Fifth or nearest level", 5},
            { "Sixth or nearest level", 6}
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

            Resampling.ItemsSource = ResamplingValues;
            Resampling.SelectedIndex = Math.Max(0, ResamplingValues.IndexOf(GlobaleSettings.ImageResamplingValue)); // Default to 72 DPI
            Resampling.IsEnabled = GlobaleSettings.ImageOption == ImageOptionChoice.Enum.Resample;

            NotesPosition.ItemsSource = NotesPositionChoices.Keys;
            NotesPosition.SelectedIndex = NotesPositionChoices.Values.ToList().IndexOf(GlobaleSettings.FootnotesPosition);

            NotesNumbering.ItemsSource = NotesNumberingChoices.Keys;
            NotesNumbering.SelectedIndex = NotesNumberingChoices.Values.ToList().IndexOf(GlobaleSettings.FootnotesNumbering);

            NotesLevel.ItemsSource = NotesLevelChoices.Keys;
            NotesLevel.SelectedIndex = Math.Max(0, GlobaleSettings.FootnotesLevel);

            StartNumber.IsEnabled = GlobaleSettings.FootnotesNumbering != FootnotesNumberingChoice.Enum.None;
            StartNumber.Text = GlobaleSettings.FootnotesStartValue.ToString().PadLeft(3,'_');
            NumberPrefix.Text = GlobaleSettings.FootnotesNumberingPrefix;
            NotesTextPrefix.Text = GlobaleSettings.FootnotesNumberingSuffix;

            UseDAISYPipelineApp.IsChecked = ConverterHelper.PipelineAppIsInstalled() && GlobaleSettings.UseDAISYPipelineApp;
            PreferredVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
            TTSEngines.IsEnabled = UseDAISYPipelineApp.IsChecked == true;
            BrowseVoices.IsEnabled = UseDAISYPipelineApp.IsChecked == true;

            TTSConfigFile.Text = GlobaleSettings.TTSConfigFile;
            TTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked == false;
            BrowseTTSConfigFile.IsEnabled = UseDAISYPipelineApp.IsChecked == false;
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
                GlobaleSettings.ImageResamplingValue = Resampling.SelectedIndex >= 0 ? ResamplingValues[Resampling.SelectedIndex] : 72; // Default to 72 DPI if not selected
                GlobaleSettings.FootnotesPosition = NotesPositionChoices.Values.ToList()[NotesPosition.SelectedIndex];
                GlobaleSettings.FootnotesLevel = NotesLevelChoices.Values.ToList()[NotesLevel.SelectedIndex];
                GlobaleSettings.FootnotesNumbering = NotesNumberingChoices.Values.ToList()[NotesNumbering.SelectedIndex];
                GlobaleSettings.FootnotesStartValue = int.Parse(StartNumber.Text.Trim().Replace("_",""));
                GlobaleSettings.FootnotesNumberingPrefix = NumberPrefix.Text;
                GlobaleSettings.FootnotesNumberingSuffix = NotesTextPrefix.Text;
                GlobaleSettings.UseDAISYPipelineApp = UseDAISYPipelineApp.IsChecked == true;
                GlobaleSettings.TTSConfigFile = TTSConfigFile.Text.Trim();
                // Save
                GlobaleSettings.Save();
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
            } else {
                GlobaleSettings.UseDAISYPipelineApp = true;
                PreferredVoices.IsEnabled = true;
                TTSEngines.IsEnabled = true;
                BrowseVoices.IsEnabled = true;
                TTSConfigFile.IsEnabled = false;
                BrowseTTSConfigFile.IsEnabled = false;
            }
        }

        private void PreferredVoices_Click(object sender, RoutedEventArgs e)
        {
            try {
                ((AppRunner)AppRunner.GetInstance()).PreferredVoices();
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not open preferred voices : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void TTSEngines_Click(object sender, RoutedEventArgs e)
        {
            try {
                ((AppRunner)AppRunner.GetInstance()).TTSEngines();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not open TTS engines settings : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void BrowseVoices_Click(object sender, RoutedEventArgs e)
        {

            try {
                ((AppRunner)AppRunner.GetInstance()).BrowseVoices();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "Could not browse voices : " + ex.Message, MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private void BrowseResultsFolder_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
