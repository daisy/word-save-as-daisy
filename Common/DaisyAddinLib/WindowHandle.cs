using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Threading;
using System.Reflection;
using System.Diagnostics;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion
{

    ///<summary>
    /// This class retrieves the IWin32Window from the current active Word window.
    /// This could be used to set the parent for Windows Forms and MessageBoxes.
    ///</summary>
    ///<example>
    /// WordWin32Window parentWindow = new WordWin32Window (ThisAddIn.WordApplication);   
    /// MessageBox.Show (parentWindow, "This MessageBox doesn't go behind Word !!!", "Attention !", MessageBoxButtons.Ok , MessageBoxIcon.Question );
    ///</example>
    public class WordWin32Window : IWin32Window
    {

        ///<summary>
        /// Returns a Handle to the Desktop-Window
        ///</summary>
        ///<returns>The handle</returns>
        [DllImport("user32")]
        private static extern IntPtr GetDesktopWindow();

        ///<summary>
        /// Finds a Window-Handle by its name or classname
        ///</summary>
        ///<param name="hwndParent">The Parent Windows-Handle</param>
        ///<param name="hwndChildAfter">Start the search after the given Windows-Handle</param>
        ///<param name="lpszClass">The classname to search for</param>
        ///<param name="lpszWindow">The window caption to search for</param>
        ///<returns><returns the windows-Handle or IntPtr.Zero</returns>
        [DllImport("user32")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        ///<summary>
        /// Gets the title of a Window.
        ///</summary>
        ///<param name="hWnd">The Windows Handle.</param>
        ///<param name="lpString">A pointer to a StringBuilder-Object.</param>
        ///<param name="nMaxCount">Max num of chars to get.</param>
        [DllImport("user32", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        ///<summary>
        /// Returns a ProcessId of the process where the window handle is accosiated to.
        ///</summary>
        ///<param name="hWnd">The window handle</param>
        ///<param name="lpdwProcessId">[out] the processId</param>
        ///<returns></returns>
        [DllImport("user32")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out IntPtr lpdwProcessId);

        #region IWin32Window Members

        ///<summary>
        /// This holds the window handle for the found Window.
        ///</summary>
        IntPtr _windowHandle = IntPtr.Zero;

        ///<summary>
        /// The <b>Handle</b> of the Word WindowObject.
        ///</summary>
        public IntPtr Handle
        {
            get { return _windowHandle; }
        }

        #endregion


        public WordWin32Window()
        {

            // get the current ProcessId
            IntPtr myProcessId = new IntPtr(System.Diagnostics.Process.GetCurrentProcess().Id);

            // Find all Word Windows-Handles

            // Find the first instance
            IntPtr hWnd = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "OpusApp", null);

            while (hWnd != IntPtr.Zero)
            {
                // Get the processId of the Word-Application instance
                IntPtr hProcess = IntPtr.Zero;
                GetWindowThreadProcessId(hWnd, out hProcess);

                if (myProcessId == hProcess)
                {
                    // gotcha
                    _windowHandle = hWnd;
                    break;
                }

                // Next - starting after the last found window
                hWnd = FindWindowEx(IntPtr.Zero, hWnd, "OpusApp", null);
            }

            Debug.Assert(_windowHandle != IntPtr.Zero, "The Word-Application parent window could not be found");
        }


        ///<summary>
        /// The <b>WordWin32Window</b> class could be used to get the parent IWin32Window for Windows.Forms and MessageBoxes.
        ///</summary>
        ///<param name="wordApp">The Word ApplicationObject.</param>
        public WordWin32Window(object wordApp)
        {

            // get the old caption of the Word Application
            string oldCaption = GetProperty<string>(wordApp, "Caption");

            // did it worked ??
            if (!string.IsNullOrEmpty(oldCaption))
            {

                // unique identifier
                string newCaption = Guid.NewGuid().ToString();

                // set a new caption to the Word Window
                if (SetProperty(wordApp, "Caption", newCaption))
                {

                    try
                    {
                        // Find all Word Windows-Handles
                        // Find the first instance
                        IntPtr hDesktop = GetDesktopWindow();
                        IntPtr hWnd = FindWindowEx(hDesktop, IntPtr.Zero, "OpusApp", null);
                        StringBuilder windowText = new StringBuilder(255);
                        while (hWnd != IntPtr.Zero)
                        {
                            // Get the window-caption
                            int a = GetWindowText(hWnd, windowText, 255);
                            if (windowText.ToString().EndsWith(newCaption))
                            {
                                // gotcha
                                _windowHandle = hWnd;
                                break;
                            }
                            // Next - starting after the last found window
                            hWnd = FindWindowEx(hDesktop, hWnd, "OpusApp", null);
                        }
                    }
                    finally
                    {
                        // Reset the caption
                        wordApp.GetType().InvokeMember("Caption", BindingFlags.PutDispProperty, null, wordApp, new object[] { oldCaption });
                    }
                }
            }

            Trace.Assert(_windowHandle != IntPtr.Zero, "The Word-Application parent window could not be found");
        }


        private T GetProperty<T>(object comObject, string propertyName)
        {
            try
            {
                return (T)comObject.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, comObject, null);
            }
            catch (System.Exception ex)
            {
                return default(T);
            }
        }

        private bool SetProperty(object comObject, string propertyName, object value)
        {
            try
            {
                comObject.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, comObject, new object[] { value });
                return true;
            }
            catch (System.Exception ex)
            {
                return false;
            }
        }


    }


}