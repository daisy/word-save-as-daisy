using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using static Daisy.SaveAsDAISY.Conversion.ConverterSettings;

namespace Daisy.SaveAsDAISY.WPF
{
    /// <summary>
    /// Logique d'interaction pour Settings.xaml
    /// </summary>
    public partial class Settings : Window
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

        

        public Settings()
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
    }
}
