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
using Daisy.SaveAsDAISY.Conversion;
using stdole;


namespace Daisy.SaveAsDAISY
{
	/// <summary>
    /// Addin UI resources accessor class (labels and logo)
    /// </summary>
    public class AddinResources
    {
        private ChainResourceManager resourceManager;

        /// <summary>
        /// Constructor
        /// </summary>
        public AddinResources()
        {
            this.resourceManager = new ChainResourceManager();
            // Add a default resource managers (for common labels)
            this.resourceManager.Add(
                new System.Resources.ResourceManager(
                    "DaisyAddinLib.resources.Labels",
                    Assembly.GetExecutingAssembly()
                )
            );
        }


        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager CustomResourceManager
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


        /// <summary>
        /// Returns the logo of the application.
        /// TODO : could be moved to addin helper class
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
            return AddinResources.ConvertImage.Convert(image);
        }

        /// <summary>
        /// Create a random temporary folder 
        /// TODO To be moved to the addinhelper class
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        /// <param name="targetExtension">The target extension</param>
        /// <returns></returns>
        public string GetTempPath(string fileName, string targetExtension)
        {
            string folderName;
            string path;
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
