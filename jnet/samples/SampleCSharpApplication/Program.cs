using org.daisy.jnet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SampleApplication {
    class Program {
        static void Main(string[] args) {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string workingDir = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path)) + Path.DirectorySeparatorChar;

            // Instantiate the JNI interface assembly
            JavaNativeInterface jni = new JavaNativeInterface();
            List<string> options = new List<string>();
            
            // Setting the class path to the jar that containes the classes to use
            options.Add("-Djava.class.path="+ workingDir + "target\\SampleJavaApplication-0.0.1-SNAPSHOT.jar");
            // If your jar need other jars as dependencies, you may need to add them in the classpath :
            // + ";" + workingDir + "target\\dependency.jar"); 

            // Load a new JVM
            jni.LoadVM(options, false);
            try {
                IntPtr SampleApplicationClass = jni.GetJavaClass("org/daisy/jnet/SampleApplication");
                IntPtr SampleApplicationObject = jni.NewObject(SampleApplicationClass);
                string testString = jni.CallMethod<string>(SampleApplicationClass, SampleApplicationObject, "getTestString", "()Ljava/lang/String;");
                Console.WriteLine(testString);

                SampleApplicationObject = jni.NewObject(SampleApplicationClass, "(Ljava/lang/String;)V", "I'm a string sent from C#");
                testString = jni.CallMethod<string>(SampleApplicationClass, SampleApplicationObject, "getTestString", "()Ljava/lang/String;");
                Console.WriteLine(testString);

                jni.CallVoidMethod(SampleApplicationClass, IntPtr.Zero, "main", "()V");
                jni.CallVoidMethod(SampleApplicationClass, IntPtr.Zero, "main", "([Ljava/lang/String;)V", new object[]{ new string[] { } });


            } catch (Exception e) {
                Console.WriteLine(e.ToString());
            }

        }
    }
}
