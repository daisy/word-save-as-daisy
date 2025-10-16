using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace org.daisy.jnet {

    /// <summary>
    /// Utilitary class to adapt java iterable
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class JavaIterable<T> : IEnumerable<T> {
        private IntPtr javaIterableClass, javaIterableObject;
        private JavaNativeInterface jni;
        public JavaIterable(JavaNativeInterface jni, IntPtr javaIterableObject) {
            this.jni = jni;
            this.javaIterableClass = jni.JNIEnvironment.GetObjectClass(javaIterableObject);
            this.javaIterableObject = javaIterableObject;
        }

        public IEnumerator<T> GetEnumerator() {
            IntPtr iteratorObject = jni.CallMethod<IntPtr>(javaIterableClass, javaIterableObject, "iterator", "()Ljava/util/Iterator;");
            IntPtr iteratorClass = jni.JNIEnvironment.GetObjectClass(iteratorObject);
            while (jni.CallMethod<bool>(iteratorClass, iteratorObject,"hasNext","()Z")) {
                yield return jni.CallMethod<T>(iteratorClass, iteratorObject, "next", "()Ljava/lang/Object;");
            }
        }

        IEnumerator IEnumerable.GetEnumerator() {
            throw new NotImplementedException();
        }
    }
}
