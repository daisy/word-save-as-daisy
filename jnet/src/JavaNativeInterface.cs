////////////////////////////////////////////////////////////////////////////////////////////////  
//  An excellent resource for the JNI library is on the website 
// Java Native Interface: Programmer's Guide and Specification by Sheng Liang 
//  http://docs.oracle.com/javase/7/docs/technotes/guides/jni/
// for a list of all the functions 
// http://download.oracle.com/javase/6/docs/technotes/guides/jni/spec/functions.html
////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace org.daisy.jnet {

    /// <summary>
    /// Main interface to simplify JNI usage
    /// </summary>
    public unsafe class JavaNativeInterface : IDisposable {
        private static string __jvmDllPath = "";
        private static string __javaVersion = "";


        // Possible registry keys that indicate the current installed JRE on windows
        private static readonly string[] JAVA_REGISTRY_KEYS = {
            @"HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\JRE",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Development Kit",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\JDK"
        };

        // Classe pointer found by class names (cashing pointers to limit jni call)
       // private readonly Dictionary<string, IntPtr> usedClasses = new Dictionary<string, IntPtr>();

        //private readonly Dictionary<KeyValuePair<IntPtr,string> , IntPtr> usedMethods = new Dictionary<KeyValuePair<IntPtr, string>, IntPtr>();

        // Currently instantiated object, for futur disposal
        //private readonly List<IntPtr> usedObjects = new List<IntPtr>();

        private JavaVM jvm;
        private JNIEnv env;

        private bool loaded = false;

        /// <summary>
        /// Original JNI environment
        /// </summary>
        public JNIEnv JNIEnvironment {
            get => env;
        }


        /// <summary>
        /// Load the jni and connect to the jvm <br/>
        /// If no custom jre path is given, search the jvm library in the given order :<br/>
        /// - Near the current assembly<br/>
        /// - Under the JAVA_HOME folder<br/>
        /// - for windows OS, also search in the user registry
        /// </summary>
        /// <param name="options">List of options to be passed while loading the jvm</param>
        /// <param name="customJREPath">Path of the jvm lib or of the directory to search jvm lib file</param>
        /// <param name="AddToExistingJVM"></param>
        /// <param name="targetVersion">Version of the jni targeted, default to JNI 10</param>
        /// <exception cref="Exception"></exception>
        public JavaNativeInterface(List<string> options = null, string customJREPath = "", bool AddToExistingJVM = false, JNIVersion targetVersion = JNIVersion.JNI_VERSION_10) {
            // os specific jvm lib names, default for windows 
            string libraryName = "jvm.dll";
            string libFolder = "bin";
            

            if (customJREPath.Length > 0) {
                if (customJREPath.EndsWith(libraryName) && File.Exists(customJREPath))
                {
                    __jvmDllPath = customJREPath;
                }
                else
                {
                    if (!Directory.Exists(customJREPath))
                    {
                        Console.Error.WriteLine($"The jre path provided was not found : {customJREPath}");
                        Console.Error.WriteLine("Falling back to standard JRE search, starting near the current assembly.");
                    }
                    string[] searchResult = Directory.GetFiles(customJREPath, libraryName, SearchOption.AllDirectories);

                    if (searchResult.Length > 0)
                    {
                        __jvmDllPath = searchResult[0];
                    }
                    else
                    {
                        Console.Error.WriteLine($"the jre path provided ({customJREPath}) does not contains a {libraryName} file");
                        Console.Error.WriteLine("Falling back to standard JRE search, starting near the current assembly.");
                    }
                }

            } 
            
            if (__jvmDllPath.Length == 0) {
                // Search for a java runtime near the current assembly
                string codeBase = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(codeBase) + Path.DirectorySeparatorChar;
                
                string[] searchResult = Directory.GetFiles(assemblyDir, libraryName, SearchOption.AllDirectories);

                if (searchResult.Length > 0) {
                    __jvmDllPath = searchResult[0];
                } else {
                    // Search a JAVA_HOME
                    string envJavaHome = Environment.GetEnvironmentVariable("JAVA_HOME");
                    if (envJavaHome != null && envJavaHome.Length > 0) {
                        searchResult = new string[] { Path.Combine(envJavaHome, libFolder, "server", libraryName) };
                        if (File.Exists(searchResult[0])) {
                            __jvmDllPath = searchResult[0];
                        }
                    } else {
                        // try to access registry
                        foreach (string key in JAVA_REGISTRY_KEYS) {
                            string javaVersion = (string)Microsoft.Win32.Registry.GetValue(key, "CurrentVersion", null);
                            if (javaVersion == "") continue;
                            else {
                                __javaVersion = javaVersion;
                                string javaKey = Path.Combine(key, javaVersion);
                                string javaHomeKey = (string)Microsoft.Win32.Registry.GetValue(javaKey, "JavaHome", null);
                                if (javaHomeKey == "") continue;
                                else {
                                    searchResult = Directory.GetFiles(javaHomeKey, libraryName, SearchOption.AllDirectories);
                                    if (searchResult.Length > 0) {
                                        __jvmDllPath = searchResult[0];
                                        break;
                                    } else continue;

                                }
                            }
                        }
                    }
                }
            }
            if (__jvmDllPath.Length == 0)
            {
                throw new Exception(
                    "No Java runtime was found near the program or in your system.\r\n" +
                    "Please contact your IT Administrator, or install yourself a JDK or a JRE and set the JAVA_HOME environment variable."
                );
            }

#if DEBUG
            Console.WriteLine("Using " + __jvmDllPath);
#endif
            JavaVM.loadAssembly(__jvmDllPath);
            loaded = true;
            LoadVM(options ?? new List<string>(), AddToExistingJVM, targetVersion);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="options"></param>
        /// <param name="AddToExistingJVM"></param>
        /// <param name="targetVersion">Specified the version</param>
        public void LoadVM(List<string> options, bool AddToExistingJVM = false, JNIVersion targetVersion = JNIVersion.JNI_VERSION_10) {

            // Set the directory to the location of the JVM.dll. 
            // This will ensure that the API call JNI_CreateJavaVM will work
            // TODO : maybe remove this after changing DLLImport by Assembly dynamic loading
            // This probably break pathes to ressources
            // Directory.SetCurrentDirectory(Path.GetDirectoryName(JavaNativeInterface.__jvmDllPath));

            var args = new JavaVMInitArgs();

            args.version = (int)targetVersion;

            args.ignoreUnrecognized = JavaVM.BooleanToByte(true); // True

            if (options.Count > 0) {
                args.nOptions = options.Count;
                var opt = new JavaVMOption[options.Count];
                int i = 0;
                foreach (string option in options) {
                    opt[i++].optionString = Marshal.StringToHGlobalAnsi(option);
                }
                fixed (JavaVMOption* a = &opt[0]) {
                    // prevents the garbage collector from relocating the opt variable as this is used in unmanaged code that the gc does not know about
                    args.options = a;
                }
            }

            if (!AddToExistingJVM) {
                IntPtr environment;
                IntPtr javaVirtualMachine;
                try
                {
                    int result = JavaVM.JNI_CreateJavaVM(out javaVirtualMachine, out environment, &args);
                    if (result == JNIReturnValue.JNI_EEXIST)
                    {
                        AttachToCurrentJVM(args);
                    }
                    else if (result != JNIReturnValue.JNI_OK)
                    {
                        throw new Exception("Cannot create JVM " + result.ToString());
                    }
                    else
                    {
                        jvm = new JavaVM(javaVirtualMachine);
                        env = new JNIEnv(environment);
                    }
                } catch (Exception e)
                {
                    throw e;
                }
            } else AttachToCurrentJVM(args);
        }

        public void deleteVM() {
            if (jvm != null) {
                int res = jvm.DestroyJavaVM();
                if (res != JNIReturnValue.JNI_OK) {
                    throw new Exception("DestroyJavaVM failed (" + res.ToString() + ")");
                }
                jvm = null;
                env = null;
            }
        }

        private void AttachToCurrentJVM(JavaVMInitArgs args) {
            // This is only required if you want to reuse the same instance of the JVM
            // This is especially useful if you are using JNI in a webservice. see page 89 of the
            // Java Native Interface: Programmer's Guide and Specification by Sheng Liang
            
            int nVMs;

            IntPtr javaVirtualMachine;
            int res = JavaVM.JNI_GetCreatedJavaVMs(out javaVirtualMachine, 1, out nVMs);
            if (res != JNIReturnValue.JNI_OK) {
                throw new Exception("JNI_GetCreatedJavaVMs failed (" + res.ToString() + ")");
            }
            if (nVMs > 0) {
                jvm = new JavaVM(javaVirtualMachine);
                res = jvm.AttachCurrentThread(out env, args);
                if (res != JNIReturnValue.JNI_OK) {
                    throw new Exception("AttachCurrentThread failed (" + res.ToString() + ")");
                }
            }
            
        }


        public string JavaVersion() {
            int? majorVersion = env?.GetMajorVersion();
            int? minorVersion = env?.GetMinorVersion();
            return majorVersion.ToString() + "." + minorVersion.ToString();
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~JavaNativeInterface() {
            Dispose();
        }

        #region Creating java class and objects

        /// <summary>
        /// Retrieve a java class reference
        /// </summary>
        /// <param name="ClassName">class path in packages using / as separator, like "org/apache/whateverClassYouNeed" </param>
        /// <returns>The class reference pointer</returns>
        public IntPtr GetJavaClass(string ClassName) {
            lock (this)
            {
                IntPtr? javaClass;
                //if (!usedClasses.ContainsKey(ClassName))
                //{
                //    try
                //    {
                //        javaClass = env?.FindClass(ClassName);
                //        if (javaClass == null)
                //        {
                //            throw new Exception(string.Format("{0} was not found", ClassName));
                //        }
                //        else usedClasses.Add(ClassName, javaClass.Value);
                //    }
                //    catch (Exception)
                //    {
                //        throw;
                //    }
                //}
                //else
                //{
                //    javaClass = usedClasses[ClassName];
                //}
                try {
                    javaClass = env?.FindClass(ClassName);
                    if (javaClass == null) {
                        throw new Exception(string.Format("{0} was not found", ClassName));
                    } //else usedClasses.Add(ClassName, javaClass.Value);
                }
                catch (Exception) {
                    throw;
                }
                return javaClass.Value;
            }
        }

        /// <summary>
        /// Create a new object for a given java class pointer
        /// </summary>
        /// <param name="javaClass">Java class pointer, retrieved by GetJavaClass method</param>
        /// <param name="signature">constructor signature in JNI format</param>
        /// <param name="args">Constructor parameters values. <br/>
        /// <b>beware that passing empty arrays as parameter might require to do a "hard declaration" of the args value, 
        /// like new object[]{ new string[]{} }  </b> </param>
        /// <returns>A java object pointer</returns>
        /// <exception cref="Exception">throws back any Java exception found during the method call</exception>
        public IntPtr NewObject(IntPtr javaClass, string signature = "()V", params object[] args) {
            lock (this)
            {
                try {
                    
                    IntPtr? methodId = env?.GetMethodID(javaClass, "<init>", signature);
                    if (methodId == null)
                    {
                        throw new Exception($"the constructor with {signature} does not exist on this class");
                    }
                    IntPtr? javaObject = env?.NewObject(javaClass, methodId.Value, ParseParameters(javaClass, signature, args));
                    if (javaObject == null) {
                        throw new Exception("The object returned was null");
                     }
                    // Store for disposal
                    //usedObjects.Add(javaObject.Value);
                    return javaObject.Value;
                }
                catch (Exception e)
                {
                    StringBuilder exceptionMessage = new StringBuilder($"An error occured while instanciating an object of the class referenced by pointer {javaClass} " +
                        $"using the constructor with signature {signature}  " +
                        $"and parameters [");
                    foreach (object arg in args) {
                        exceptionMessage.Append(arg.ToString());
                        exceptionMessage.Append(", ");
                    }
                    exceptionMessage.Append("]");
                    throw new Exception(exceptionMessage.ToString(), e);
                }

            } 
        }



        /// <summary>
        /// Helper function to create a new java wrapper object from a C# object based on the input object type. <br/>
        /// The following object types are mapped to the indicated java wrapper :<br/>
        /// - bool : java/lang/Boolean <br/>
        /// - short : java/lang/Short <br/>
        /// - int : java/lang/Integer <br/>
        /// - long : java/lang/Long <br/>
        /// - float : java/lang/Float <br/>
        /// - double : java/lang/Double <br/>
        /// - string : java/lang/String <br/>
        /// - List : java/util/ArrayList <br/>
        /// - Dictionnary(string,object) : java/util/HashMap
        /// </summary>
        /// <param name="obj">Object of expected type : bool, short, int, float, double, string, List or Dictionnary (with string key)</param>
        /// <returns>IntPtr.Zero if the object is an instanced of a non-supported type, else a java pointer to a wrapper object.</returns>
        public IntPtr NewJavaWrapperObject(object obj)
        {
            if(obj is string)
            {
                IntPtr javaClass = GetJavaClass("java/lang/String");
                return NewObject(javaClass, "(Ljava/lang/String;)V", (string)obj);
            }
            else if (obj is long)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Long");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(J)Ljava/lang/Long;",
                    (long)obj
                );
            }
            else if(obj is int)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Integer");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(I)Ljava/lang/Integer;",
                    (int)obj
                );
            }
            else if (obj is short)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Short");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(S)Ljava/lang/Short;",
                    (short)obj
                );
            }
            else if (obj is bool)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Boolean");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(Z)Ljava/lang/Boolean;",
                    (bool)obj
                );
            }
            else if (obj is float)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Float");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(F)Ljava/lang/Float;",
                    (float)obj
                );
            }
            else if (obj is double)
            {
                IntPtr javaClass = GetJavaClass("java/lang/Double");
                return CallMethod<IntPtr>(
                    javaClass,
                    IntPtr.Zero,
                    "valueOf",
                    "(D)Ljava/lang/Double;",
                    (double)obj
                );
            }
            else if (obj is IList) // Generic list interface
            {
                return NewArrayListFrom((IList)obj);
            }
            else if (obj is IDictionary) // Generic dictionnary interface
            {
                return NewHashMapFrom((IDictionary)obj);
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// Make a java ArrayList from a Csharp list of item
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <returns></returns>
        public IntPtr NewArrayListFrom(IList list)
        {
            IntPtr javaListClass = GetJavaClass("java/util/ArrayList");

            IntPtr newList = NewObject(javaListClass);
            foreach (object item in list)
            {
                bool tempBool = CallMethod<bool>(
                    javaListClass,
                    newList,
                    "add",
                    "(Ljava/lang/Object;)Z",
                    NewJavaWrapperObject(item)
                );
            }

            return newList;
        }

        public IntPtr NewHashMapFrom(IDictionary map)
        {
            IntPtr javaMapClass = GetJavaClass("java/util/HashMap");
            IntPtr newMap = NewObject(javaMapClass);
            // Not optimal but it seems
            // there is not a generic KeyValuePair interface available
            // But i can go through keys and retrieve values from them like this
            foreach (object key in map.Keys)
            {
                object value = map[key];
                CallMethod<IntPtr>(
                    javaMapClass,
                    newMap,
                    "put",
                    "(Ljava/lang/Object;Ljava/lang/Object;)Ljava/lang/Object;",
                    NewJavaWrapperObject(key),
                    NewJavaWrapperObject(value)
                );
            }

            return newMap;
        }

        /// <summary>
        /// Retrieve values from a java array pointer
        /// </summary>
        /// <typeparam name="T">expected type of values in the array</typeparam>
        /// <param name="array">Pointer of the array within a class or and object</param>
        /// <returns>the array of values of a default empty array if no type found</returns>
        private T[] GetArray<T>(IntPtr array) {
            lock (this)
            {
                try
                {
                    if (env == null)
                    {
                        throw new Exception("JNI was not initialised before calling GetArray");
                    }

                    if (typeof(T) == typeof(byte) || typeof(T) == typeof(sbyte))
                    {
                        // Get the byte array values 
                        return (T[])(object)env.GetByteArray(array);
                    }
                    else if (typeof(T) == typeof(bool))
                    {
                        // Get the boolean array values 
                        return (T[])(object)env.GetBooleanArray(array);
                    }
                    else if (typeof(T) == typeof(char))
                    {
                        // Get the char Field 
                        return (T[])(object)env.GetCharArray(array);
                    }
                    else if (typeof(T) == typeof(short))
                    {
                        // Get the short Field 
                        return (T[])(object)env.GetShortArray(array);
                    }
                    else if (typeof(T) == typeof(int))
                    {
                        // Get the int array values               
                        return (T[])(object)env.GetIntArray(array);
                    }
                    else if (typeof(T) == typeof(long))
                    {
                        // Get the long array values 
                        return (T[])(object)env.GetLongArray(array);
                    }
                    else if (typeof(T) == typeof(float))
                    {
                        // Get the float array values 
                        return (T[])(object)env.GetFloatArray(array);
                    }
                    else if (typeof(T) == typeof(double))
                    {
                        // Get the double array values 
                        return (T[])(object)env.GetDoubleArray(array);
                    }
                    else if (typeof(T) == typeof(string))
                    {
                        // Get the string array values
                        IntPtr[] objArray = env.GetObjectArray(array);
                        string[] res = new string[objArray.Length];

                        for (int i = 0; i < objArray.Length; i++)
                        {
                            res[i] = env.JStringToString(objArray[i]);
                        }
                        return (T[])(object)res;
                    }
                    else if (typeof(T) == typeof(IntPtr))
                    {
                        // Get the object array values
                        return (T[])(object)env.GetObjectArray(array);
                    }
                    else
                    {
                        return default;
                    }


                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">C# Type of the field to return</typeparam>
        /// <param name="javaClass">The class pointer</param>
        /// <param name="javaObjectPointer">An object pointer, can be null or IntPtr.Zero for static fields</param>
        /// <param name="FieldName">Name of the field in the class</param>
        /// <param name="sig">Java signature of the field</param>
        /// <returns></returns>
        public T GetField<T>(IntPtr javaClass, IntPtr? javaObjectPointer, string FieldName, string sig) {
            lock (this)
            {
                try
                {
                    if (env == null)
                    {
                        throw new Exception(string.Format("JNI was not initialised before calling GetField for {0} with signature {1}", FieldName, sig));
                    }

                    IntPtr javaObject = javaObjectPointer == null ? IntPtr.Zero : javaObjectPointer.Value;

                    bool isStatic = javaObject == IntPtr.Zero;

                    IntPtr FieldID = isStatic ?
                        env.GetStaticFieldID(javaClass, FieldName, sig) :
                        env.GetFieldID(javaClass, FieldName, sig);
                    if (typeof(T) == typeof(byte) || typeof(T) == typeof(sbyte))
                    {
                        // Get the byte Field 
                        sbyte res = isStatic ?
                            env.GetStaticByteField(javaClass, FieldID) :
                            env.GetByteField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(sbyte[])
                           || typeof(T) == typeof(byte[]))
                    { // For now, use sbyte for both
                      // Possibly add a warning in debug if byte is requested
                      // Get the byte Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        sbyte[] res = GetArray<sbyte>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(bool))
                    {
                        // Get the boolean Field 
                        bool res = isStatic ?
                            env.GetStaticBooleanField(javaClass, FieldID) :
                            env.GetBooleanField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(bool[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        bool[] res = GetArray<bool>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(char))
                    {
                        // Get the char Field 
                        char res = isStatic ?
                            env.GetStaticCharField(javaClass, FieldID) :
                            env.GetCharField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(char[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        char[] res = GetArray<char>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(short))
                    {
                        // Get the short Field 
                        short res = isStatic ?
                            env.GetStaticShortField(javaClass, FieldID) :
                            env.GetShortField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(short[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        short[] res = GetArray<short>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(int))
                    {
                        // Get the int Field               
                        int res = isStatic ?
                            env.GetStaticIntField(javaClass, FieldID) :
                            env.GetIntField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(int[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        int[] res = GetArray<int>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(long))
                    {
                        // Get the long Field 
                        long res = isStatic ?
                            env.GetStaticLongField(javaClass, FieldID) :
                            env.GetLongField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(long[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        long[] res = GetArray<long>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(float))
                    {
                        // Get the float Field 
                        float res = isStatic ?
                            env.GetStaticFloatField(javaClass, FieldID) :
                            env.GetFloatField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(float[]))
                    {
                        // Get the float array values
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        float[] res = GetArray<float>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(double))
                    {
                        // Get the double Field 
                        double res = isStatic ?
                            env.GetStaticDoubleField(javaClass, FieldID) :
                            env.GetDoubleField(javaObject, FieldID);
                        return (T)(object)res; // need to fix this
                    }
                    else if (typeof(T) == typeof(double[]))
                    {
                        // Get the double array field values
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        double[] res = GetArray<double>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(string))
                    {
                        // Get the string Field 
                        IntPtr jstr = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);

                        string res = env.JStringToString(jstr);
                        env.DeleteLocalRef(jstr);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(string[]))
                    {
                        // Get the string array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        string[] res = GetArray<string>(jobj); ;
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(IntPtr))
                    {
                        // Get the object Field
                        IntPtr res = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(IntPtr[]))
                    {
                        // Get the int array Field
                        IntPtr jobj = isStatic ?
                            env.GetStaticObjectField(javaClass, FieldID) :
                            env.GetObjectField(javaObject, FieldID);
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        IntPtr[] res = GetArray<IntPtr>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    return default(T);


                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        #endregion


        #region Methods calls

        /// <summary>
        /// Call a void-returning method from an object and/or java class. <br/>
        /// (This additionnal function is needed due to void type not being usable as type in generics.)
        /// </summary>
        /// <param name="javaClass">pointer of the java class</param>
        /// <param name="javaObject">pointer of the object to used as caller. <br/>
        ///     if IntPtr.Zero, method is considered static.
        /// </param>
        /// <param name="methodName">method name</param>
        /// <param name="sig">Method JNI signature (see javap -s javaClass.class )</param>
        /// <param name="param">parameters to used with </param>
        /// <exception cref="Exception">throws back any Java exception found during the method call</exception>
        public void CallVoidMethod(IntPtr javaClass, IntPtr javaObject, string methodName, string sig, params object[] param) {
            lock (this)
            {
                try
                {
                    if (env == null)
                    {
                        throw new Exception(string.Format("JNI was not initialised before trying to call method {0} with signature {1}", methodName, sig));
                    }

                    if (javaObject != IntPtr.Zero)
                    {
                        IntPtr methodId = env.GetMethodID(javaClass, methodName, sig);
                        env.CallVoidMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                    }
                    else
                    {
                        IntPtr methodId = env.GetStaticMethodID(javaClass, methodName, sig);
                        env.CallStaticVoidMethod(javaClass, methodId, ParseParameters(javaClass, sig, param));
                    }
                }

                catch (Exception)
                {
                    throw;
                }
            }
            
        }

        /// <summary>
        /// Call a method returning a T value from an object and/or a java class. <br/>
        /// If a correct object pointer is provided (that is not null nor IntPtr.Zero), it tries to call the method from the object.
        /// If null or zero pointer is provided for object, it will try to launch a corresponding static method.<br/>
        /// <b>If the method has no return value (returns void), you must use CallVoidMethod.</b>
        /// </summary>
        /// <typeparam name="T">
        ///     Method return type equivalent in c# (expects a non-void csharp raw type, including arrays like int[], or IntPtr)<br/>
        ///     <b>Beware that byte type correspond to "sbyte" in java</b> 
        /// </typeparam>
        /// <param name="javaClass">Java class pointer</param>
        /// <param name="javaObject">Java Object pointer</param>
        /// <param name="methodName">Name of the method to call</param>
        /// <param name="sig">Method's JNI signature</param>
        /// <param name="args">Constructor parameters values. <br/>
        /// <b>beware that passing empty arrays as parameter might require to do a "hard declaration" of the args value, 
        /// like new object[]{ new string[]{} }  </b> </param>
        /// <returns></returns>
        /// <exception cref="Exception">throws back any Java exception found during the method call</exception>
        public T CallMethod<T>(IntPtr javaClass, IntPtr? javaObjectPointer, string methodName, string sig, params object[] param) {
            lock (this)
            {
                try
                {
                    if (env == null)
                    {
                        throw new Exception(string.Format("JNI was not initialised before calling CallMethod for {0} with signature {1}", methodName, sig));
                    }

                    IntPtr javaObject = javaObjectPointer == null ? IntPtr.Zero : javaObjectPointer.Value;
                    bool isStatic = javaObject == IntPtr.Zero;

                    IntPtr methodId = isStatic ?
                        env.GetStaticMethodID(javaClass, methodName, sig) :
                        env.GetMethodID(javaClass, methodName, sig);
                    if (typeof(T) == typeof(byte) || typeof(T) == typeof(sbyte))
                    {
                        // Call the byte method 
                        sbyte res = isStatic ?
                            env.CallStaticByteMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallByteMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(byte[]) || typeof(T) == typeof(sbyte[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        sbyte[] res = GetArray<sbyte>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(bool))
                    {
                        // Call the boolean method 
                        bool res = isStatic ?
                            env.CallStaticBooleanMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallBooleanMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(bool[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        bool[] res = GetArray<bool>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(char))
                    {
                        // Call the char method 
                        char res = isStatic ?
                            env.CallStaticCharMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallCharMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(char[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        char[] res = GetArray<char>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(short))
                    {
                        // Call the short method 
                        short res = isStatic ?
                            env.CallStaticShortMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallShortMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(short[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        short[] res = GetArray<short>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(int))
                    {
                        // Call the int method               
                        int res = isStatic ?
                            env.CallStaticIntMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallIntMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(int[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        int[] res = GetArray<int>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(long))
                    {
                        // Call the long method 
                        long res = isStatic ?
                            env.CallStaticLongMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallLongMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(long[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        long[] res = GetArray<long>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(float))
                    {
                        // Call the float method 
                        float res = isStatic ?
                            env.CallStaticFloatMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallFloatMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(float[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        float[] res = GetArray<float>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(double))
                    {
                        // Call the double method 
                        double res = isStatic ?
                            env.CallStaticDoubleMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallDoubleMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res; // need to fix this
                    }
                    else if (typeof(T) == typeof(double[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        double[] res = GetArray<double>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(string))
                    {
                        // Call the string method 
                        IntPtr jstr = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));

                        string res = env.JStringToString(jstr);
                        env.DeleteLocalRef(jstr);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(string[]))
                    {
                        // Call the string array method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }

                        IntPtr[] objArray = env.GetObjectArray(jobj);
                        string[] res = new string[objArray.Length];

                        for (int i = 0; i < objArray.Length; i++)
                        {
                            res[i] = env.JStringToString(objArray[i]);
                        }

                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(IntPtr))
                    {
                        // Call the object method and deal with whatever comes back in the call code 
                        IntPtr res = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        return (T)(object)res;
                    }
                    else if (typeof(T) == typeof(IntPtr[]))
                    {
                        // Call the byte method
                        IntPtr jobj = isStatic ?
                            env.CallStaticObjectMethod(javaClass, methodId, ParseParameters(javaClass, sig, param)) :
                            env.CallObjectMethod(javaObject, methodId, ParseParameters(javaClass, sig, param));
                        if (jobj == IntPtr.Zero)
                        {
                            return default(T);
                        }
                        IntPtr[] res = GetArray<IntPtr>(jobj);
                        env.DeleteLocalRef(jobj);
                        return (T)(object)res;
                    }
                    return default(T);

                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        #endregion




        protected virtual void Dispose(bool disposing) {
            lock (this)
            {
                // free native resources if there are any.
                //foreach (KeyValuePair<string, IntPtr> javaClass in usedClasses)
                //{
                //    if (javaClass.Value != IntPtr.Zero)
                //    {
                //        env?.DeleteGlobalRef(javaClass.Value);
                //        usedClasses[javaClass.Key] = IntPtr.Zero;
                //    }
                //}
                //usedClasses.Clear();
                //for (int i = 0, end = usedObjects.Count; i < end; ++i)
                //{
                //    IntPtr javaObject = usedObjects[i];
                //    if (javaObject != IntPtr.Zero)
                //    {
                //        env?.DeleteLocalRef(javaObject);
                //        usedObjects[i] = IntPtr.Zero;
                //    }
                //}
                //usedObjects.Clear();

                if (disposing)
                {
                    // free managed resources
                    if (jvm != null)
                    {
                        jvm.Dispose();
                        jvm = null;
                    }

                    if (env != null)
                    {
                        env.Dispose();
                        env = null;
                        
                    }
                }
            }
            
        }

        #region Parameters parsing utilities
        /// <summary>
        /// Parse parameters and signature to create the parameter array.
        /// </summary>
        /// <param name="javaClass"></param>
        /// <param name="sig"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        private JValue[] ParseParameters(IntPtr javaClass, string sig, params object[] param) {
            
            JValue[] retval = new JValue[param.Length];

            int startIndex = sig.IndexOf('(') + 1;

            for (int i = 0; i < param.Length; i++) {
                string paramSig = "";
                if (sig.Substring(startIndex, 1) == "[") {
                    paramSig = sig.Substring(startIndex, 1);
                    startIndex += 1;
                }
                if (sig.Substring(startIndex, 1) == "L") {
                    paramSig += sig.Substring(startIndex, sig.IndexOf(';', startIndex) - startIndex);
                    startIndex++; // skip past ;
                } else {
                    paramSig += sig.Substring(startIndex, 1);
                }

                startIndex += (paramSig.Length - (paramSig.IndexOf("[", StringComparison.Ordinal) + 1));
                // changed 
                if (param[i] is string)
                {
                    // also adding Object test for generics
                    if (!(paramSig.Equals("Ljava/lang/String") || paramSig.Equals("Ljava/lang/Object")))
                    {
                        throw new Exception("Signature (" + paramSig + ") does not match parameter value (" + param[i].GetType().ToString() + ").");
                    }
                    if (env == null)
                    {
                        throw new Exception(string.Format("JNI was not initialised before trying to create a java string"));
                    }
                    lock (env)
                    {
                        retval[i] = new JValue() { L = env.NewString(param[i].ToString(), param[i].ToString().Length) };
                    }
                }
                else if (param[i] == null)
                {
                    retval[i] = new JValue(); // Just leave as default value
                }
                else if (paramSig.StartsWith("["))
                {
                    retval[i] = ProcessArrayType(javaClass, paramSig, param[i]);
                }
                else if (param[i] is IntPtr)
                {   // java object pointer
                    retval[i] = new JValue() { L = (IntPtr)param[i] };
                }else {
                    retval[i] = new JValue();

                    FieldInfo paramField = retval[i].GetType().GetFields(BindingFlags.Public | BindingFlags.Instance).AsQueryable().FirstOrDefault(a => a.Name.ToUpper().Equals(paramSig));
                    if ((paramField != null)
                        && (
                            (param[i].GetType() == paramField.FieldType)
                            || (
                                (paramField.FieldType == typeof(byte) || paramField.FieldType == typeof(bool))
                                && (param[i] is byte || param[i] is bool)
                            )
                        )
                    ) {
                        paramField.SetValueDirect(
                            __makeref(retval[i]),
                             (paramField.FieldType == typeof(byte) || paramField.FieldType == typeof(bool))   // this is an undocumented feature to set struct fields via reflection
                                ? JavaVM.BooleanToByte((bool)param[i])
                                : param[i]
                        );
                    } else throw new Exception("Signature (" + paramSig + ") does not match parameter value (" + param[i].GetType().ToString() + ").");
                }
            }
            return retval;
        }


        private JValue ProcessArrayType(IntPtr javaClass, string paramSig, object param) {
            if (env == null)
            {
                throw new Exception(string.Format("JNI was not initialised before trying to process an array type"));
            }
            
            IntPtr arrPointer;
            if (paramSig.Equals("[I"))
                arrPointer = env.NewIntArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[J"))
                arrPointer = env.NewLongArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[C"))
                arrPointer = env.NewCharArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[B"))
                arrPointer = env.NewByteArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[S"))
                arrPointer = env.NewShortArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[D"))
                arrPointer = env.NewDoubleArray(((Array)param).Length, javaClass);
            else if (paramSig.Equals("[F"))
                arrPointer = env.NewFloatArray(((Array)param).Length, javaClass);
            else if (paramSig.Contains("[Ljava/lang/String"))
            {
                IntPtr jclass = env.FindClass("Ljava/lang/String;");
                try
                {
                    arrPointer = env.NewObjectArray(((Array)param).Length, jclass, IntPtr.Zero);
                }
                finally
                {
                    env.DeleteLocalRef(jclass);
                }

            }
            else if (paramSig.Contains("[Ljava/lang/"))
                arrPointer = env.NewObjectArray(((Array)param).Length, javaClass, (IntPtr)param);
            else
            {
                throw new Exception("Signature (" + paramSig + ") does not match parameter value (" +
                                    param.GetType().ToString() + "). All arrays types should be defined as objects because I do not have enough time to defines every possible array type");
            }

            if (paramSig.Contains("[Ljava/lang/"))
            {
                for (int j = 0; j < ((Array)param).Length; j++)
                {
                    object obj = ((Array)param).GetValue(j);

                    if (paramSig.Contains("[Ljava/lang/String"))
                    {
                        IntPtr str = env.NewString(obj.ToString(), obj.ToString().Length);
                        env.SetObjectArrayElement(arrPointer, j, str);
                    }
                    else
                        env.SetObjectArrayElement(arrPointer, j, (IntPtr)obj);
                }
            }
            else
                env.PackPrimitiveArray<int>((int[])param, arrPointer);

            return new JValue() { L = arrPointer };
            
        }
        #endregion
    }
}
