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
using System.Collections;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using System.IO.Packaging;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// An XmlUrlResolver for zip packaged files. 
    /// This resolver is mostly invoked when using the XSL document() function.
	/// </summary>
	public class ZipResolver : XmlUrlResolver, IDisposable
    {
        private const string ZIP_URI_SCHEME = "zip";
        private const string ZIP_URI_HOST = "localhost";
        private Package archive;
		private Hashtable entries;
		
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="filename">Absolute path to the Zip package</param>
		public ZipResolver(String filename)
		{
			archive = Package.Open(filename);
			// initialize hash table of Daisy resource file names
			entries = new Hashtable();
		}

        public Package Archive
        {
            get {return archive;}
        }

        public void Dispose() {
            Dispose(true);
        }
        protected void Dispose(bool disposing) {
            if (disposing) {
                // Dispose managed resources
                if (archive != null) {
                    archive.Close();
                }
            }
        }


        /* Function to give URI of a file*/
        public override Uri ResolveUri(Uri baseUri, String relativeUri)
		{
            if (baseUri == null)
            {
                if (relativeUri == null || relativeUri.Length == 0 || relativeUri.Contains("://"))
                {
                    return null;
                }
                else
                {
                    Uri uri = new Uri(ZIP_URI_SCHEME + "://" + ZIP_URI_HOST + "/" + relativeUri);
                    if (!entries.ContainsKey(uri.AbsoluteUri))
                    {
                        entries.Add(uri.AbsoluteUri, relativeUri);
                    }
                    return uri;
                }
            }
            else
            {
                return base.ResolveUri(baseUri, relativeUri);
            }
		}

        /* Function to get Stream of a file using its URI*/
        public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
		{
			try {
				
				Stream stream = null;
				
				if (entries.Contains(absoluteUri.AbsoluteUri))
				{
                    PackagePart archivepart;
					                  
                    string aparturi;
                    aparturi = "/" + (string)entries[absoluteUri.AbsoluteUri];

                    Uri archiveparturi = new Uri(aparturi, UriKind.Relative);
                    archivepart = archive.GetPart(archiveparturi);
                    stream = archivepart.GetStream();
				}
				
				if (stream == null)
				{
					stream = new MemoryStream();
				}
                // Cannot have a 0 byte xml document !
                else if (stream.Length == 0)
                { 
                    throw new IOException(entries[absoluteUri.AbsoluteUri] + " is 0 length");
                }
                return stream;
			} 
			catch (Exception e) 
			{
                // failsafe on a dummy xml document
                string s;
                s = e.Message;

				return EmbeddedResourceResolver.GetSharedResource("source.xml");
			}
		}
		
		
	}
			
	
}
