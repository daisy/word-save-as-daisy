using System;
using System.Collections.Generic;
using System.Text;
using SpeechLib;
using System.Threading;


namespace SimpleCommandLineTTS
{
    class Program
    {
        static bool debug = false;
        static void Main(string[] args)
        {
            int inputCounter = 0;
            string input = "";
            string filename = "";
            SpVoice voice = new SpVoice();


            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.UTF8;
    
            DEBUG("SimpleCommandLine#Main(string[]): Using input encoding:  " + Console.InputEncoding.BodyName);
            DEBUG("SimpleCommandLine#Main(string[]): Using output encoding: " + Console.OutputEncoding.BodyName);

            while ((input = Console.ReadLine()) != null)
            {
                DEBUG("SimpleCommandLine#Main(string[]): Read: " + input);
                input = input.Trim();
                if (input.Length == 0)
                {
                    return;
                }

                inputCounter++;
                if ((inputCounter % 2) == 0)
                {
                    if (say(filename, input, voice))
                    {
                        Console.Out.WriteLine("OK");
                        Console.Out.Flush();
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    filename = input;
                }
            }
        }

        static bool say(string filename, string sent, SpVoice voice)
        {
            SpFileStream fileStream = null;
            bool success = true;
            try
            {
                SpeechVoiceSpeakFlags flags = SpeechVoiceSpeakFlags.SVSFIsXML;
                //SpeechVoiceSpeakFlags flags = SpeechVoiceSpeakFlags.SVSFIsNotXML;
                //SpVoice voice = new SpVoice();
                SpeechStreamFileMode fileMode = SpeechStreamFileMode.SSFMCreateForWrite;

                fileStream = new SpFileStream();
                fileStream.Open(filename, fileMode, false);

                // audio output
                /*
                voice.Speak(sent, flags);
                voice.WaitUntilDone(Timeout.Infinite);
                */
                // file output
                voice.AudioOutputStream = fileStream;
                voice.Speak(sent, flags);
                voice.WaitUntilDone(Timeout.Infinite);
            }
            catch (Exception error)
            {
                success = false;
                Console.Error.WriteLine("Error speaking sentence: " + error);
                Console.Error.WriteLine("error.Data: " + error.Data);
                Console.Error.WriteLine("error.HelpLink: " + error.HelpLink);
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Close();
                }
            }
            return success;
        }

        private static void DEBUG(string msg)
        {
            if (debug)
            {
                Console.Error.WriteLine(msg);
            }
        }
    }
}
