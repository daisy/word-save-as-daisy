using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Pipeline.ChainedScripts;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using Daisy.SaveAsDAISY.WPF;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WPFFormsTests
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TestAboutWindows_Click(object sender, RoutedEventArgs e)
        {
            About aboutWindow = new About();
            aboutWindow.Show();
        }

        private void TestSettingsWindows_Click(object sender, RoutedEventArgs e)
        {
            SettingsForm settingsWindow = new SettingsForm();
            settingsWindow.Show();
        }

        private void TestConversionWindows_Click(object sender, RoutedEventArgs e)
        {
            Script s = new WordToEpub3(null);
            DocumentProperties testProps = new DocumentProperties(@"C:\Users\npavie\Downloads\9782213731179_JoursDeRevolte_RL_BRL.docx");
            File.Copy(testProps.InputPath, testProps.CopyPath);
            testProps.updatePropertiesFromCopy();

            ConversionParametersForm conversionForm = new ConversionParametersForm(s, testProps);
            conversionForm.ShowDialog();
        }

        private void TestMetadata_Click(object sender, RoutedEventArgs e)
        {
            Metadata metaWin = new Metadata();
            metaWin.Show();
        }
    }
}
