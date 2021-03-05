using System;
using System.Collections;
using System.Collections.Generic;

namespace Daisy.DaisyConverter.DaisyConverterLib
{
	/// <summary>
	/// Input parameters for convert operaton.
	/// </summary>
    public class OoxToDaisyParameters
    {
        public string InputFile { get; set; }
        public string TempInputFile { get; set; }
        public string TempInputA { get; set; }
        public string Version { get; set; }
        public string ControlName { get; set; }
        public ArrayList ObjectShapes { get; set; }
        public Hashtable ListMathMl { get; set; }
        public ArrayList ImageIds { get; set; }
        public ArrayList InlineShapes { get; set; }
        public ArrayList InlineIds { get; set; }
        public string MasterSubFlag { get; set; }
		public string ScriptPath { get; set; }
		public string Directory { get; set; }

		public string GetInputFileNameWithoutExtension
        {
            get
            {
                int lastSeparatorIndex = InputFile.LastIndexOf('\\');
                // Special case : onedrive documents uses https based URL format with '/' as separator
                if(lastSeparatorIndex < 0) {
                    lastSeparatorIndex = InputFile.LastIndexOf('/');
                }
                if (lastSeparatorIndex < 0) { // no path separator found
                    return InputFile.Remove(InputFile.LastIndexOf('.'));
                } else {
                    string tempInput = InputFile.Substring(lastSeparatorIndex);
                    return tempInput.Remove(tempInput.LastIndexOf('.'));
                }
            }
        }
    }
}