using System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

namespace org.daisy.jnet {

    /// <summary>
    /// JNI Version to pass to the interface.
    /// Possible values :<br/>
    /// JNI_VERSION_1_1, (possibly not supported)<br/>
    /// JNI_VERSION_1_2,<br/>
    /// JNI_VERSION_1_4,<br/>
    /// JNI_VERSION_1_6,<br/>
    /// JNI_VERSION_1_8,<br/>
    /// JNI_VERSION_9,  <br/>
    /// JNI_VERSION_10 (dafault targeted version) <br/>
    /// </summary>
    public enum JNIVersion : int {
        /// <summary>
        /// Java 1
        /// </summary>
        JNI_VERSION_1_1 = 0x00010001,
        /// <summary>
        /// Java 2
        /// </summary>
        JNI_VERSION_1_2 = 0x00010002,
        /// <summary>
        /// Java 4
        /// </summary>
        JNI_VERSION_1_4 = 0x00010004,
        /// <summary>
        /// Java 6
        /// </summary>
        JNI_VERSION_1_6 = 0x00010006,
        /// <summary>
        /// Java 8
        /// </summary>
        JNI_VERSION_1_8 = 0x00010008,
        /// <summary>
        /// Java 9
        /// </summary>
        JNI_VERSION_9 = 0x00090000,
        /// <summary>
        /// Java 10 and superior
        /// </summary>
        JNI_VERSION_10 = 0x000a0000
}

    public struct JNIBooleanValue
    {
        // JBoolean Constant
        public const byte JNI_FALSE = 0;  // False
        public const byte JNI_TRUE = 1;   // True (wouldn't it be fun to set JNI_TRUE = 0 that would give someone hrs of fun of debugging)
        
    }


    /// <summary>
    /// Release mode for ReleaseScalarArrayElements
    /// </summary>
    public struct JNIPrimitiveArrayReleaseMode {
        public const int JNI_DEFAULT = 0;
        public const int JNI_COMMIT = 1;
        public const int JNI_ABORT = 2;
    }

    /// <summary>
    /// 
    /// </summary>
    public struct JNIReturnValue
    {
        // possible return values for JNI functions
        public const int JNI_OK = 0;  // success
        public const int JNI_ERR = -1; // unknown error
        public const int JNI_EDETACHED = -2; // thread detached from the VM
        public const int JNI_EVERSION = -3; // JNI version error
        public const int JNI_ENOMEM = -4; // not enough memory 
        public const int JNI_EEXIST = -5; // VM already created 
        public const int JNI_EINVAL = -6; // invalid arguments 

        public const int JNI_ENOJava = 101; // local error if the DLL can not be found


        public static string Message(int code) {
            switch (code) {
                case JNI_OK:
                    return "Success";
                case JNI_EDETACHED:
                    return "Thread detached from the VM";
                case JNI_EVERSION:
                    return "JNI version error";
                case JNI_ENOMEM:
                    return "Not enough memory";
                case JNI_EEXIST:
                    return "VM already created";
                case JNI_EINVAL:
                    return "Invalid arguments";
                case JNI_ENOJava:
                    return "Dll related local error";
                case JNIReturnValue.JNI_ERR:
                    return "Unknown error";
                default:
                    return "Unknown error code " + code;
            }
        }

    }

    // Invocation API
    [StructLayout(LayoutKind.Sequential), NativeCppClass]
    public struct JavaVMOption
    {
        public IntPtr optionString;
        public IntPtr extraInfo;
    }

    [StructLayout(LayoutKind.Sequential), NativeCppClass]
    public unsafe struct JavaVMInitArgs
    {
        public int version;
        public int nOptions;
        public JavaVMOption* options;
        public byte ignoreUnrecognized;
    }

    public unsafe struct JavaVMAttachArgs {
        public int version;
        public IntPtr name; // char*
        public IntPtr group; // jobject
    }

    // You can't have reference types (jObject etc). So I have dropped the jObject type and replaced it with a IntPtr
    [StructLayout(LayoutKind.Explicit)]
    public struct JValue
    {
        /// <summary>
        /// boolean type in signatures
        /// </summary>
        [FieldOffset(0)]
        public byte Z;

        /// <summary>
        /// byte type in signatures
        /// </summary>
        [FieldOffset(0)]
        public byte B;

        /// <summary>
        /// char type in signatures
        /// </summary>
        [FieldOffset(0)]
        public char C;

        /// <summary>
        /// short type in signatures
        /// </summary>
        [FieldOffset(0)]
        public short S;

        /// <summary>
        /// Boolean type in signatures
        /// </summary>
        [FieldOffset(0)]
        public int I;

        /// <summary>
        /// long type in signatures
        /// </summary>
        [FieldOffset(0)]
        public long J;

        /// <summary>
        /// float type in signatures
        /// </summary>
        [FieldOffset(0)]
        public float F;

        /// <summary>
        /// double type in signatures
        /// </summary>
        [FieldOffset(0)]
        public double D;

        /// <summary>
        /// Fully qualified class of object in signatures
        /// </summary>
        [FieldOffset(0)]
        public IntPtr L;
    }

    // JDK 1.6
    public enum jobjectRefType {
        JNIInvalidRefType = 0,
        JNILocalRefType = 1,
        JNIGlobalRefType = 2,
        JNIWeakGlobalRefType = 3
    }
}



