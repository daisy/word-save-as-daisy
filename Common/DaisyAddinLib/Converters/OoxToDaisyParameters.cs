using System;
using System.Collections;
using System.Collections.Generic;

namespace Sonata.DaisyConverter.DaisyConverterLib
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
                string tempInput = InputFile.Substring(InputFile.LastIndexOf('\\'));
                return tempInput.Remove(tempInput.LastIndexOf('.'));
            }
        }
    }
}