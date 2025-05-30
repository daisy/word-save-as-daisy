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

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// An XmlUrlResolver for embedded resources.
	/// </summary>
	public class SharedXmlUrlResolver : XmlUrlResolver
    {
		private Assembly shared;
		private bool isDirectTransform;
		
	    
		public SharedXmlUrlResolver(bool isDirectTransform) 
		{
			this.shared = Assembly.GetExecutingAssembly();
            this.isDirectTransform = isDirectTransform;
		}

        public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
        {   
            if (!File.Exists(absoluteUri.LocalPath))
            {
                // look into shared embedded resources
                string resource = absoluteUri.OriginalString.Substring(absoluteUri.OriginalString.LastIndexOf('/')+1);
                string path = isDirectTransform ? this.GetType().Namespace + ".resources.Daisy2oox." + resource :
                                                this.GetType().Namespace + ".resources.oox2Daisy." + resource;

                return this.shared.GetManifestResourceStream(path);
            }
            else
            {
                return base.GetEntity(absoluteUri, role, ofObjectToReturn);
            }
        }
	
	}

}
