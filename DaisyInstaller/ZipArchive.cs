using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.IO;

namespace DaisyInstaller {
    // ZipArchive from https://www.codeproject.com/Articles/209731/Csharp-use-Zip-archives-without-external-libraries
    class ZipArchive : IDisposable {
        private object external;
        private ZipArchive() { }
        public enum CompressionMethodEnum { Stored, Deflated };
        public enum DeflateOptionEnum { Normal, Maximum, Fast, SuperFast };
        //...
        public static ZipArchive OpenOnFile(string path, FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read, FileShare share = FileShare.Read, bool streaming = false) {
            var type = typeof(System.IO.Packaging.Package).Assembly.GetType("MS.Internal.IO.Zip.ZipArchive");
            var meth = type.GetMethod("OpenOnFile", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipArchive { external = meth.Invoke(null, new object[] { path, mode, access, share, streaming }) };
        }
        public static ZipArchive OpenOnStream(Stream stream, FileMode mode = FileMode.OpenOrCreate, FileAccess access = FileAccess.ReadWrite, bool streaming = false) {
            var type = typeof(System.IO.Packaging.Package).Assembly.GetType("MS.Internal.IO.Zip.ZipArchive");
            var meth = type.GetMethod("OpenOnStream", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipArchive { external = meth.Invoke(null, new object[] { stream, mode, access, streaming }) };
        }
        public ZipFileInfo AddFile(string path, CompressionMethodEnum compmeth = CompressionMethodEnum.Deflated, DeflateOptionEnum option = DeflateOptionEnum.Normal) {
            var type = external.GetType();
            var meth = type.GetMethod("AddFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            var comp = type.Assembly.GetType("MS.Internal.IO.Zip.CompressionMethodEnum").GetField(compmeth.ToString()).GetValue(null);
            var opti = type.Assembly.GetType("MS.Internal.IO.Zip.DeflateOptionEnum").GetField(option.ToString()).GetValue(null);
            return new ZipFileInfo { external = meth.Invoke(external, new object[] { path, comp, opti }) };
        }
        public void DeleteFile(string name) {
            var meth = external.GetType().GetMethod("DeleteFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            meth.Invoke(external, new object[] { name });
        }
        public void Dispose() {
            ((IDisposable)external).Dispose();
        }
        public ZipFileInfo GetFile(string name) {
            var meth = external.GetType().GetMethod("GetFile", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return new ZipFileInfo { external = meth.Invoke(external, new object[] { name }) };
        }

        public IEnumerable<ZipFileInfo> Files {
            get {
                var meth = external.GetType().GetMethod("GetFiles", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                var coll = meth.Invoke(external, null) as System.Collections.IEnumerable; //ZipFileInfoCollection
                foreach (var p in coll) yield return new ZipFileInfo { external = p };
            }
        }
        public IEnumerable<string> FileNames {
            get { return Files.Select(p => p.Name).OrderBy(p => p); }
        }

        public struct ZipFileInfo {
            internal object external;
            private object GetProperty(string name) {
                return external.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(external, null);
            }
            public override string ToString() {
                return Name;// base.ToString();
            }
            public string Name {
                get { return (string)GetProperty("Name"); }
            }
            public DateTime LastModFileDateTime {
                get { return (DateTime)GetProperty("LastModFileDateTime"); }
            }
            public bool FolderFlag {
                get { return (bool)GetProperty("FolderFlag"); }
            }
            public bool VolumeLabelFlag {
                get { return (bool)GetProperty("VolumeLabelFlag"); }
            }
            public object CompressionMethod {
                get { return GetProperty("CompressionMethod"); }
            }
            public object DeflateOption {
                get { return GetProperty("DeflateOption"); }
            }
            public Stream GetStream(FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read) {
                var meth = external.GetType().GetMethod("GetStream", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                return (Stream)meth.Invoke(external, new object[] { mode, access });
            }
        }
    }
}
