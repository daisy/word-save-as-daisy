using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ApplicationWithIntegratedJRE {
    class Program {
        static void Main(string[] args) {
            
            // Same as the SampleCsharpApplication, but the build_jre script is called in a pre-built pass


            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string workingDir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path)) + Path.DirectorySeparatorChar;

            
            List<string> options = new List<string>();

            // Setting the class path to the jar that containes the classes to use
            options.Add("-Djava.class.path="+workingDir + "target\\SampleJavaAppWithJRE-0.0.1-SNAPSHOT.jar");
            // If your jar need other jars as dependencies, you may need to add them in the classpath :
            // + ";" + workingDir + "target\\dependency.jar"); 
            
            // Instantiate the JNI interface assembly
            JavaNativeInterface jni = new JavaNativeInterface(options);
            try {
                IntPtr SampleApplicationClass = jni.GetJavaClass("org/daisy/jnet/SampleApplication");

                IntPtr SampleApplicationObject = jni.NewObject(SampleApplicationClass);
                string testString = jni.CallMethod<string>(SampleApplicationClass, SampleApplicationObject, "getTestString", "()Ljava/lang/String;");
                Console.WriteLine(testString);

                SampleApplicationObject = jni.NewObject(SampleApplicationClass, "(Ljava/lang/String;)V", "I'm a string sent from C# 日本");
                testString = jni.CallMethod<string>(SampleApplicationClass, SampleApplicationObject, "getTestString", "()Ljava/lang/String;");
                Console.WriteLine(testString);

                //IntPtr DefaultPackageSampleClass = IntPtr.Zero;
                //IntPtr DefaultPackageSampleObject = jni.NewObject(DefaultPackageSampleClass);

                //testString = jni.CallMethod<string>(DefaultPackageSampleClass, DefaultPackageSampleObject, "getExecutablePath", "()Ljava/lang/String;");
                //Console.WriteLine(testString);


            } catch (Exception e) {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
