/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Collections;
using Daisy.DaisyConverter.DaisyConverterLib.Converters;
using stdole;


namespace Daisy.DaisyConverter.DaisyConverterLib
{
	/// <summary>
    /// Base class MS Office add-in implementations.
    /// </summary>
    public class DaisyAddinLib
    {
        private AbstractConverter converter;
        private ChainResourceManager resourceManager;
        DesignForm myForm;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="converter">An implementation of AbstractConverter</param>
        public DaisyAddinLib(AbstractConverter converter)
        {
            this.converter = converter;
            this.resourceManager = new ChainResourceManager();
            // Add a default resource managers (for common labels)
            this.resourceManager.Add(new System.Resources.ResourceManager("DaisyAddinLib.resources.Labels",
                Assembly.GetExecutingAssembly()));
        }
        
        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager
        {
            set { this.resourceManager.Add(value); }
        }


        /// <summary>
        /// Retrieve the label associated to the specified key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public string GetString(string key)
        {
            return this.resourceManager.GetString(key);
        }

        /// <summary>
        /// Resource manager
        /// </summary>
        public System.Resources.ResourceManager ResManager
        {
            get { return this.resourceManager; }
        }

        public void OoxToDaisyWithoutUI(OoxToDaisyParameters parameters, TranslationParametersBuilder parametersBuilder, string outputFilePath, string outputPipeline)
        {
            Hashtable translationParams = parametersBuilder.BuildTranslationParameters();
            ConvertToDaisy(parameters, translationParams, outputFilePath, outputPipeline, new SingleConverter(converter));
        }

        /// <summary>
        /// Funcion that launch the conversion of a "parameters.inputFile" to daisy
        /// </summary>
        /// <param name="parameters"></param>
        /// <param name="translateParams"></param>
        /// <param name="outputFilePath"></param>
        /// <param name="outputPipeline"></param>
        /// <param name="singleConverter"></param>
        public void ConvertToDaisy(OoxToDaisyParameters parameters, Hashtable translateParams, string outputFilePath, string outputPipeline, SingleConverter singleConverter)
        {
            if (parameters.MasterSubFlag == "No" || parameters.MasterSubFlag == "NoMasterSub")
            {
                //outputFilePath = AddInHelper.buttonIsSingleWordToXMLConversion(parameters.ControlName) ?
                //        outputFilePath + parameters.GetInputFileNameWithoutExtension + ".xml" :
                //        Path.Combine(outputFilePath, "convertedDocument" + ".xml");
                outputFilePath = (outputFilePath + parameters.GetInputFileNameWithoutExtension + ".xml").Replace(',','_');
                singleConverter.OoxToDaisy(
                        parameters.InputFile, 
                        outputFilePath, 
                        parameters.ListMathMl, 
                        translateParams,
                        parameters.ControlName, 
                        outputPipeline);
            }
            else if (parameters.MasterSubFlag == "Yes")
            {
                singleConverter.OoxToDaisyOwn(
                    parameters.TempInputFile,
                    parameters.TempInputA, 
                    parameters.InputFile, 
                    outputFilePath, 
                    translateParams, 
                    parameters.ControlName, 
                    parameters.ListMathMl,
                    outputPipeline);
            }
        }

        /// <summary>
        /// Star the conversion of a single word file to selected fornat (
        /// - Function which shows UI to take input from user
        /// </summary>
        public void StartSingleWordConversion(OoxToDaisyParameters parameters) {
            if (parameters.ScriptPath != null && parameters.ScriptPath.Length > 0) {
                myForm = new DesignForm(
                       parameters.ScriptPath,
                       parameters.InputFile,
                       parameters.Directory,
                       parameters.Version,
                       parameters.ControlName,
                       parameters.TempInputFile,
                       this.resourceManager,
                       parameters.MasterSubFlag);
            } else {
                myForm = new DesignForm(parameters, this.resourceManager);
            }

            int translateFlag = myForm.DoTranslate();
            if (translateFlag == 1) {
                SingleConverter singleConverter = new SingleConverterUI(
                        converter,
                        parameters.ScriptPath != null && parameters.ScriptPath.Length > 0 ? 
                            myForm.getParser : 
                            null);
                ConvertToDaisy(
                        parameters,
                        myForm.HTable,
                        parameters.ScriptPath.Length > 0 ?
                            AddInHelper.AppDataSaveAsDAISYDirectory :
                            myForm.OutputFilepath,
                        myForm.PipeOutput,
                        singleConverter);
            }
        }

		/*public void StartSingleWordConversion(OoxToDaisyParameters parameters)
        {
            myForm = new DesignForm(
                parameters.ScriptPath, 
                parameters.InputFile, 
                parameters.Directory, 
                parameters.Version, 
                parameters.ControlName, 
                parameters.TempInputFile, 
                this.resourceManager, 
                parameters.MasterSubFlag);
            int translateFlag = myForm.DoTranslate();
            
            if (translateFlag == 1)
            {
                SingleConverter singleConverter = new SingleConverterUI(converter, myForm.getParser);
                ConvertToDaisy(parameters, myForm.HTable, AddInHelper.AppDataSaveAsDAISYDirectory, myForm.PipeOutput,singleConverter);
            }
        }*/


		public void OoxToDaisyDTBookWithoutUI(OoxToDaisyParameters parameters, TranslationParametersBuilder parametersBuilder, string outputFilePath, string outputPipeline, ScriptParser parser)
		{
			Hashtable translationParams = parametersBuilder.BuildTranslationParameters();
			ConvertToDaisy(parameters, translationParams, outputFilePath, outputPipeline, new SingleConverter(converter, parser));
		}

		public bool OoxToDaisySub(String outputfilepath, ArrayList subList, String category, Hashtable table, string control, Hashtable MultipleMathMl, string output_Pipeline)
        {
            SingleConverter singleConverter = new SingleConverterUI(converter, null);
            return singleConverter.OoxToDaisySub(outputfilepath, subList,  category, table, control, MultipleMathMl, output_Pipeline);
        }

        public bool OoxToDaisySubWithoutUI(String outputfilepath, ArrayList subList, String category, Hashtable table, string control, Hashtable MultipleMathMl, string output_Pipeline)
        {
            SingleConverter singleConverter = new SingleConverter(converter, null);
            return singleConverter.OoxToDaisySub(outputfilepath, subList, category, table, control, MultipleMathMl, output_Pipeline);
        }

        /// <summary>
        /// Returns the logo of the application.
        /// </summary>
        /// <param name="iconFile">Name of FIle</param>
        /// <returns>The logo of the application.</returns>
        public stdole.IPictureDisp GetLogo(String iconFile)
        {
            Assembly asmBly = Assembly.GetExecutingAssembly();
            Stream streamIcon = null;
            foreach (string name in asmBly.GetManifestResourceNames())
            {
                if (name.EndsWith(iconFile))
                {
                    streamIcon = asmBly.GetManifestResourceStream(name);
                    break;
                }
            }
            if (streamIcon == null)
            {
                return null;
            }
            System.Drawing.Bitmap image = new System.Drawing.Bitmap(streamIcon);
            return DaisyAddinLib.ConvertImage.Convert(image);
        }

        /// <summary>
        /// Create a random temporary folder 
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        /// <param name="targetExtension">The target extension</param>
        /// <returns></returns>
        public string GetTempPath(string fileName, string targetExtension)
        {
            string folderName = null;
            string path = null;
            do
            {
                folderName = Path.GetRandomFileName();
                path = Path.Combine(Path.GetTempPath(), folderName);
            }
            while (Directory.Exists(path));

            Directory.CreateDirectory(path);
            return Path.Combine(path, Path.GetFileNameWithoutExtension(fileName) + targetExtension);
        }

        #region nested types

        /*Function which converts an image into OLE IPicture object*/
        sealed private class ConvertImage : AxHost
        {
            private ConvertImage()
                : base(null)
            {
            }
            public static IPictureDisp Convert
                (Image image)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(image);


            }
        }

        #endregion
    }
}
