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
using System.IO;
using System.Collections;
using System.Text;
using System.Reflection;
using System.Xml;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// An XmlUrlResolver for embedded resources.
    /// Looks for resources in the specified assemby first, then in the 'common' assembly.
	/// </summary>
	public class EmbeddedResourceResolver : XmlUrlResolver
    {
        public const string ASSEMBLY_URI_SCHEME = "assembly";
        public const string ASSEMBLY_URI_HOST = "localhost";
        private Assembly appAssembly;
		private Assembly defaultAssembly;
		private string appNamespace;
        private string defaultNamespace;
		private bool isDirectTransform;
        private string resourceLocation;

        // URI resolver mapper to avoid http calls that could trigger proxy errors
        private Dictionary<string, string> dtdMapper = new Dictionary<string, string>()
        {
            // dtds
            { "-//NISO//DTD dtbook 2005-3//EN", "dtbook-2005-3.dtd" },
            { "-//W3C//DTD MathML 2.0//EN", "mathml2.dtd" },
            // entities
            { "-//W3C//ENTITIES MathML 2.0 Qualified Names 1.0//EN", "mathml2-qname-1.mod" },
            { "-//W3C//ENTITIES Added Math Symbols: Arrow Relations for MathML 2.0//EN", "isoamsa.ent" },
            { "-//W3C//ENTITIES Added Math Symbols: Binary Operators for MathML 2.0//EN", "isoamsb.ent" },
            { "-//W3C//ENTITIES Added Math Symbols: Delimiters for MathML 2.0//EN", "isoamsc.ent" },
            { "-//W3C//ENTITIES Added Math Symbols: Negated Relations for MathML 2.0//EN", "isoamsn.ent" },
            { "-//W3C//ENTITIES Added Math Symbols: Ordinary for MathML 2.0//EN", "isoamso.ent" },
            { "-//W3C//ENTITIES Added Math Symbols: Relations for MathML 2.0//EN", "isoamsr.ent" },
            { "-//W3C//ENTITIES Greek Symbols for MathML 2.0//EN", "isogrk3.ent" },
            { "-//W3C//ENTITIES Math Alphabets: Fraktur for MathML 2.0//EN", "isomfrk.ent" },
            { "-//W3C//ENTITIES Math Alphabets: Open Face for MathML 2.0//EN", "isomopf.ent" },
            { "-//W3C//ENTITIES Math Alphabets: Script for MathML 2.0//EN", "isomscr.ent" },
            { "-//W3C//ENTITIES General Technical for MathML 2.0//EN", "isotech.ent" },
            { "-//W3C//ENTITIES Box and Line Drawing for MathML 2.0//EN", "isobox.ent" },
            { "-//W3C//ENTITIES Russian Cyrillic for MathML 2.0//EN", "isocyr1.ent" },
            { "-//W3C//ENTITIES Non-Russian Cyrillic for MathML 2.0//EN", "isocyr2.ent" },
            { "-//W3C//ENTITIES Diacritical Marks for MathML 2.0//EN", "isodia.ent" },
            { "-//W3C//ENTITIES Added Latin 1 for MathML 2.0//EN", "isolat1.ent" },
            { "-//W3C//ENTITIES Added Latin 2 for MathML 2.0//EN", "isolat2.ent" },
            { "-//W3C//ENTITIES Numeric and Special Graphic for MathML 2.0//EN", "isonum.ent" },
            { "-//W3C//ENTITIES Publishing for MathML 2.0//EN", "isopub.ent" },
            { "-//W3C//ENTITIES Extra for MathML 2.0//EN", "mmlextra.ent" },
            { "-//W3C//ENTITIES Aliases for MathML 2.0//EN", "mmlalias.ent" },
        };
		
		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="assembly">The assembly embedding resources</param>
		/// <param name="prefix">The prefix used to locate resources within the assembly</param>
        /// <param name="isDirectTransform">Is the direction of conversion direct</param>
		public EmbeddedResourceResolver(Assembly appAssembly, string appNamespace, bool isDirectTransform) :
            this(appAssembly, appNamespace)
		{
            this.isDirectTransform = isDirectTransform;
		}

        public EmbeddedResourceResolver(Assembly appAssembly, string appNamespace, string resourceLocation, bool isDirectTransform) :
            this (appAssembly, appNamespace, isDirectTransform)
        {
            this.resourceLocation = resourceLocation;
        }

        public EmbeddedResourceResolver(Assembly appAssembly, string appNamespace)
        {
            this.appAssembly = appAssembly;
            this.appNamespace = appNamespace;
            this.defaultNamespace = this.GetType().Namespace;
            this.defaultAssembly = Assembly.GetExecutingAssembly();
        }

        /*Function to set true of false*/
        public bool IsDirectTransform
        {
            set { this.isDirectTransform = value; }
        }

		/* Function to give URI of a file*/
        public override Uri ResolveUri(Uri baseUri, String relativeUri)
		{
			if (baseUri == null && !relativeUri.Contains("://"))
            {
                if (File.Exists(relativeUri)) // pointing to an existing path on system and not a relative
                {
                    return base.ResolveUri(baseUri, relativeUri);
                } else return new Uri(ASSEMBLY_URI_SCHEME + "://" + ASSEMBLY_URI_HOST + "/" + relativeUri);
            }
            else if (relativeUri.EndsWith(".ent", StringComparison.OrdinalIgnoreCase)
                || relativeUri.EndsWith(".dtd", StringComparison.OrdinalIgnoreCase)
                || relativeUri.EndsWith(".mod", StringComparison.OrdinalIgnoreCase)
                || relativeUri.EndsWith(".css", StringComparison.OrdinalIgnoreCase)
            )
            {
                string file = Path.GetFileName(relativeUri);
                return new Uri(ASSEMBLY_URI_SCHEME + "://" + ASSEMBLY_URI_HOST + "/" + file);
            } else if (dtdMapper.ContainsKey(relativeUri))
            {
                return new Uri(ASSEMBLY_URI_SCHEME + "://" + ASSEMBLY_URI_HOST + "/" + dtdMapper[relativeUri]);
            }
            else
            {
                return base.ResolveUri(baseUri, relativeUri);
            }
		}
		
		/* Function to get Stream of a file using its URI*/
        public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
		{
            if (ASSEMBLY_URI_SCHEME.Equals(absoluteUri.Scheme))
            {
                string location = null;
                string resourceLocation = absoluteUri.OriginalString.Remove(0, ASSEMBLY_URI_SCHEME.Length + ASSEMBLY_URI_HOST.Length + 4).Replace("/", ".");

                if (this.resourceLocation == null)
                {
                    location = ".resources.";
                    switch (Path.GetExtension(resourceLocation)) {
                        case ".ent":
                        case ".mod":
                            location = location + "entities.";
                            break;
                        case ".xsl":
                            location = location + "xslts.";
                            break;
                        case ".dtd":
                            location = location + "dtds.";
                            break;
                        case ".css":
                            location = location + "css.";
                            break;
                    }
                }
                else
                {
                    location = this.resourceLocation;
                }
                
                string resource = location + resourceLocation;

                Stream stream = this.appAssembly.GetManifestResourceStream(this.appNamespace + resource);
                if (stream != null)
                {
                    return stream;
                }
                // fallback to the shared embedded resources
                else
                {
                    stream = this.defaultAssembly.GetManifestResourceStream(this.defaultNamespace+resource);
                    if (stream != null) {
                        return stream;
                    } else throw new Exception("Internal error : could not find file " + resource + " - " + this.appNamespace + resource + " - " + this.defaultNamespace + resource);
                }
            }

            // use default behaviour instead
            return base.GetEntity(absoluteUri, role, ofObjectToReturn);
		}
	
		/*Function to get Data Stream of a file*/
        public Stream GetInnerStream(String fileName) 
        {
            string location = null;
            if (this.resourceLocation == null)
            {
                location = ".resources.";
                switch (Path.GetExtension(fileName)) {
                    case ".ent":
                    case ".mod":
                        location = location + "entities.";
                        break;
                    case ".xsl":
                        location = location + "xslts.";
                        break;
                    case ".dtd":
                        location = location + "dtds.";
                        break;
                    case ".css":
                        location = location + "css.";
                        break;
                }
            }
            else
            {
                location = this.resourceLocation;
            }
            
            string path = this.appNamespace + location;
            string[] temp = this.appAssembly.GetManifestResourceNames();
            return this.appAssembly.GetManifestResourceStream(path + fileName);
		}
		

		public static Stream GetSharedResource(string filename) 
        {
			Assembly a = Assembly.GetExecutingAssembly();
			return a.GetManifestResourceStream(a.GetName()+".resources."+filename);
		}
	}

}
