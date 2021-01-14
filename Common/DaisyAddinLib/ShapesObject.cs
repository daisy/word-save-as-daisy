using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

namespace Daisy.DaisyConverter.DaisyConverterLib
{
    public enum CLIPFORMAT : int
    {
        CF_TEXT = 1,
        CF_BITMAP = 2,
        CF_METAFILEPICT = 3,
        CF_SYLK = 4,
        CF_DIF = 5,
        CF_TIFF = 6,
        CF_OEMTEXT = 7,
        CF_DIB = 8,
        CF_PALETTE = 9,
        CF_PENDATA = 10,
        CF_RIFF = 11,
        CF_WAVE = 12,
        CF_UNICODETEXT = 13,
        CF_ENHMETAFILE = 14,
        CF_HDROP = 15,
        CF_LOCALE = 16,
        CF_MAX = 17,
        CF_OWNERDISPLAY = 0x80,
        CF_DSPTEXT = 0x81,
        CF_DSPBITMAP = 0x82,
        CF_DSPMETAFILEPICT = 0x83,
        CF_DSPENHMETAFILE = 0x8E,
    }

    public class User32
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool CloseClipboard();
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EmptyClipboard();


        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern int FormatMessage(int dwFlags, string lpSource, int dwMessageId, int dwLanguageId,
                                               StringBuilder lpBuffer, int nSize, string[] Arguments);

        public static string GetLastErrorMessage() {
            StringBuilder strLastErrorMessage = new StringBuilder(255);
            int ret2 = Marshal.GetLastWin32Error();
            const int dwFlags = 4096;
            FormatMessage(dwFlags, null, ret2, 0, strLastErrorMessage, strLastErrorMessage.Capacity, null);
            return "Error code " + ret2.ToString() + " : " + strLastErrorMessage.ToString();
        }
    }

    public class ClipboardDataException : Exception {

        public ClipboardDataException(string message) : base(message) { }


    }

    public class ClipboardAccessException : Exception {

        public ClipboardAccessException(string message) : base(message) { }


    }

    /// <summary>
    /// System clipboard method to use instead of the "Clipboard" csharp tool
    /// (In our test, the Clipboard calls are intercepted by Office
    /// and Word display a "save clipboard data" request on close
    /// </summary>
    public abstract class ClipboardEx
    {
        public static System.Drawing.Imaging.Metafile GetEMF(IntPtr hWnd)
        {
            try
            {
                /* Attempt to open the Clipboard. */
                //if (OpenClipboard(GetClipboardOwner()))
                if (User32.OpenClipboard(hWnd))
                {
                    /* Check the Clipboard data format. */
                    if (User32.IsClipboardFormatAvailable((uint)CLIPFORMAT.CF_ENHMETAFILE))
                    {
                        /* Get the pointer to the data. */
                        IntPtr ptr = User32.GetClipboardData((uint)CLIPFORMAT.CF_ENHMETAFILE);
                        if (!ptr.Equals(IntPtr.Zero))
                        {
                            /* Return the Metafile. */
                            return new Metafile(ptr, true);
                        }
                        else
                        {
                            throw new ClipboardDataException("Error extracting CF_ENHMETAFILE from clipboard.\r\n- System message: \""+User32.GetLastErrorMessage() + "\""); 
                        }
                    }
                    else
                    {
                        throw new ClipboardDataException("CF_ENHMETAFILE is not available in clipboard.\r\n- System message: \"" + User32.GetLastErrorMessage() + "\"");
                    }
                }
                else
                {
                    throw new ClipboardAccessException("Error opening clipboard.\r\n- System message: \"" + User32.GetLastErrorMessage() + "\"");
                }
            }
            catch (System.Exception e) // ensures finally is executed before propagation
            {
                throw e;
            }
            finally
            {
                /* Important to close the Clipboard. */
                //User32.EmptyClipboard();
                User32.CloseClipboard();
            }
        }
    }
}
