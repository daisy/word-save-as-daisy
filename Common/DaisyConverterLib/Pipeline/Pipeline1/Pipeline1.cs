using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;

namespace Daisy.SaveAsDAISY.Conversion
{
    /// <summary>
    /// Conversion lib embedded pipeline 1 management class
    /// </summary>
    public class Pipeline1
    {
        public static string InstallationPath {
            get { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms"; }
        }

        private Dictionary<string, FileInfo> m_ScriptsInfo = new Dictionary<string, FileInfo>();
        private Dictionary<string, Script> m_Scripts = new Dictionary<string, Script>();

        #region Singleton initialisation

        private static Pipeline1 instance = null;
        // for thread safety
        private static readonly object padlock = new object();


        public static Pipeline1 Instance {
            get {
                lock (padlock) {
                    if (instance == null) {
                        instance = new Pipeline1();
                    }
                    return instance;
                }
            }
        }


        private Pipeline1(string ScriptsDirectory) {
            string[] ScriptsFilePaths = Directory.GetFiles(ScriptsDirectory, "*.taskScript", SearchOption.TopDirectoryOnly);
            foreach (string s in ScriptsFilePaths) {
                FileInfo ScriptFileInfo = new FileInfo(s);
                m_ScriptsInfo.Add(
                    ScriptFileInfo.Name.Replace(ScriptFileInfo.Extension, ""),
                    ScriptFileInfo
                );

            }
        }

        /// <summary>
        /// Default pipeline 1 installation when an addin is deployed
        /// </summary>
        private Pipeline1() : this(Path.Combine(Pipeline1.InstallationPath, @"scripts")) {
        }
        #endregion

        public void AddScriptsDirectory(string ScriptsDirectory) {
            string[] ScriptsFilePaths = Directory.GetFiles(ScriptsDirectory, "*.taskScript", SearchOption.TopDirectoryOnly);
            foreach (string s in ScriptsFilePaths) {
                FileInfo ScriptFileInfo = new FileInfo(s);
                m_ScriptsInfo.Add(
                    ScriptFileInfo.Name.Replace(ScriptFileInfo.Extension, ""),
                    ScriptFileInfo
               );

            }
        }

        public Dictionary<string, FileInfo> ScriptsInfo
        {
            get { return m_ScriptsInfo; }
        }

        public Script getScript(string key) {
            if (!m_Scripts.ContainsKey(key)) {
                if (!m_ScriptsInfo.ContainsKey(key)) throw new Exception("Call to a non-existing pipeline 1 script: " + key);
                m_Scripts.Add(
                    key,
                    new Pipeline1Script(m_ScriptsInfo[key].FullName)
                );
                
            }
            return m_Scripts[key];
        }


    }
}
