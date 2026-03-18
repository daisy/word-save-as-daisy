using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ConversionApp
{
    /// <summary>
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static Dictionary<string, string> ParseCommandLineArgs(string[] args, out string scriptName)
        {
            var options = new Dictionary<string, string>();
            scriptName = args[1];

            for (int i = 2; i < args.Length; i++) // Start from 2 to skip the executable name an script name
            {
                if (args[i].StartsWith("--"))
                {
                    string key = args[i].Substring(2);
                    if (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                    {
                        options[key] = args[i + 1];
                        i++;
                    }
                    else
                    {
                        options[key] = "true";
                    }
                }
            }

            return options;
        }

        [STAThread]
        public static int Main()
        {
            var application = new App();
            application.InitializeComponent();

            

            string[] args = Environment.GetCommandLineArgs();
            if (args.Length < 2)
            {
                MessageBox.Show("Usage: ConversionApp.exe <scriptName> [--option value]...");
                return 1;
            }
            string scriptOrCommand;
            Dictionary<string, string> options = ParseCommandLineArgs(args, out scriptOrCommand);

            switch (scriptOrCommand.ToLower())
            {
                case "scripts":
                    // export available scripts descriptors
                    break;
                case "datatypes":
                    // export available datatypes
                    break;
                case "properties":
                    // opens a properties editor, that would allow users to edit properties of the pipeline
                    // if the properties file exists, it should be loaded when starting the JVM
                    break;
                default:
                    ConversionWindow window = new ConversionWindow(scriptOrCommand, options);
                    application.MainWindow = window;
                    window.Show();
                    break;
            }
            
            return application.Run();

        }
    }
}
