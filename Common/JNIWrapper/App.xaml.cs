using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace org.daisy.jniwrapper
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
                
                if (args[i].StartsWith("--")) // script options are expected to be in the form --option value or --option (for boolean flags)
                {
                    string key = args[i].Substring(2);
                    if (i + 1 < args.Length && !(args[i + 1].StartsWith("--") || args[i + 1].StartsWith("-D")))
                    {
                        options[key] = args[i + 1];
                        i++;
                    }
                    else
                    {
                        options[key] = "true";
                    }
                } else if (args[i].StartsWith("-D")) // properties to be passed to the jni runtime
                {
                    string key = args[i];
                    if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
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
            
            Window main = null;

            string[] args = Environment.GetCommandLineArgs();
            if (args.Length < 2)
            {
                MessageBox.Show("Usage: JNIWrapper.exe <scriptName> [--option value] [-Dorg.daisy.property value]...");
                return 1;
            }
            string scriptOrCommand;
            Dictionary<string, string> options = ParseCommandLineArgs(args, out scriptOrCommand);
            // resplit command line options in properties (starting with -D) and regular options,
            // properties will be passed to the jni runtime and regular options will be passed to the script
            Dictionary<string, string> properties = options.Where(kv => kv.Key.StartsWith("-D")).ToDictionary(kv => kv.Key, kv => kv.Value);
            options = options.Where(kv => !kv.Key.StartsWith("-D")).ToDictionary(kv => kv.Key, kv => kv.Value);
            JNITaskRunner.Command command = JNITaskRunner.Command.None;
            switch (scriptOrCommand.ToLower())
            {
                case "scripts": // List of available scripts
                    command = JNITaskRunner.Command.Scripts;
                    break;
                case "script": // definition of a script
                    command = JNITaskRunner.Command.ScriptDetails;
                    break;
                case "datatypes": // List of datatype
                    // Export datatypes list to an output file
                    command = JNITaskRunner.Command.Datatypes;
                    break;
                case "datatype": // Precise definition of a datatype
                    // Export datatypes list to an output file
                    command = JNITaskRunner.Command.DatatypeDetails;
                    break;
                case "descriptors":
                    command = JNITaskRunner.Command.Descriptors;
                    break;
                case "settable-properties":
                    command = JNITaskRunner.Command.SettableProperties;
                    //string outputDir = options.ContainsKey("output") ? options["output"] : Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    //if(!Directory.Exists(outputDir))
                    //{
                    //    MessageBox.Show($"Output directory '{outputDir}' does not exist.");
                    //    return 2;
                    //}
                    //// no need for a window, we can run the generation directly and exit

                    //return JNITaskRunner.ExecuteCommand(
                    //    JNITaskRunner.Command.Descriptors, options, default, 
                    //    (info) => {Console.Out.WriteLine(info);},
                    //    null, 
                    //    (error) => {Console.Error.WriteLine(error);},
                    //    properties
                    //).Result;
                    break;
            }
            if(command != JNITaskRunner.Command.None)
            {
                // It's a command, run it and exit
                var code = JNITaskRunner.ExecuteCommand(command, options, default,
                        (info) => { Console.Out.WriteLine(info); },
                        null,
                        (error) => { Console.Error.WriteLine(error); },
                        properties
                ).Result;
                return code;
            } else  // (command == JNITaskRunner.Command.None)
            {
                // Not a command, assume it's a script name and open the window to run it
                main = new ConversionWindow(scriptOrCommand, options, properties);
                var application = new App();
                application.InitializeComponent();
                application.MainWindow = main;
                main.Show();
                int testCode = application.Run();
                return testCode;
            }

        }
    }
}
