using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Collections;

using MSword = Microsoft.Office.Interop.Word;
using IConnectDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using ConnectFORMATETC = System.Runtime.InteropServices.ComTypes.FORMATETC;
using ConnectSTGMEDIUM = System.Runtime.InteropServices.ComTypes.STGMEDIUM;
using COMException = System.Runtime.InteropServices.COMException;
using TYMED = System.Runtime.InteropServices.ComTypes.TYMED;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing;
using Microsoft.Office.Core;
using System.Windows.Input;

namespace Daisy.SaveAsDAISY.Addins.Word2007 {

    /// <summary>
    /// Document preprocessor
    /// </summary>
    public class DocumentPreprocessor : IDocumentPreprocessor {

        /// <summary>
        /// Saves the meta file. (source : https://keestalkstech.com/2016/06/rasterizing-emf-files-png-net-csharp/)
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="destination">The destination.</param>
        /// <param name="scale">The scale. Default value is 4.</param>
        /// <param name="backgroundColor">Color of the background.</param>
        /// <param name="format">The format. Default is PNG.</param>
        /// <param name="parameters">The parameters.</param>
        public static void SaveMetaFile(
            Stream source,
            Stream destination,
            float scale = 1f,
            Color? backgroundColor = null,
            ImageFormat format = null,
            EncoderParameters parameters = null)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (destination == null)
            {
                throw new ArgumentNullException(nameof(destination));
            }

            using (var img = new Metafile(source))
            {
                var f = format ?? ImageFormat.Png;

                //Determine default background color. 
                //Not all formats support transparency. 
                if (backgroundColor == null)
                {
                    var transparentFormats = new ImageFormat[] { ImageFormat.Gif, ImageFormat.Png, ImageFormat.Wmf, ImageFormat.Emf };
                    var isTransparentFormat = transparentFormats.Contains(f);

                    backgroundColor = isTransparentFormat ? Color.Transparent : Color.White;
                }

                //header contains DPI information
                var header = img.GetMetafileHeader();

                //calculate the width and height based on the scale
                //and the respective DPI
                var width = (int)Math.Round((scale * img.Width / header.DpiX * 100), 0, MidpointRounding.ToEven);
                var height = (int)Math.Round((scale * img.Height / header.DpiY * 100), 0, MidpointRounding.ToEven);

                using (var bitmap = new Bitmap(width, height))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bitmap))
                    {
                        //fills the background
                        g.Clear(backgroundColor.Value);

                        //reuse the width and height to draw the image
                        //in 100% of the square of the bitmap
                        g.DrawImage(img, 0, 0, bitmap.Width, bitmap.Height);
                    }
                    // crop image
                    int xmin = bitmap.Width - 1, xmax = 0;
                    // search min and max x
                    for (int y = 0; y < bitmap.Height; ++y)
                    {
                        for (int x = 0; x <= xmin; ++x)
                        {
                            if (bitmap.GetPixel(x, y).ToArgb() != backgroundColor.Value.ToArgb())
                            {
                                xmin = Math.Min(xmin, x);
                            }
                        }

                        for (int x = bitmap.Width - 1; x >= xmax; --x)
                        {
                            if (bitmap.GetPixel(x, y).ToArgb() != backgroundColor.Value.ToArgb())
                            {
                                xmax = Math.Max(xmax, x);
                            }
                        }
                    }
                    // search min y
                    int ymin = bitmap.Height - 1, ymax = 0;
                    for (int x = 0; x < bitmap.Width; ++x)
                    {
                        for (int y = 0; y <= ymin; ++y)
                        {
                            if (bitmap.GetPixel(x, y).ToArgb() != backgroundColor.Value.ToArgb())
                            {
                                ymin = Math.Min(ymin, y);
                            }
                        }
                        for (int y = bitmap.Height - 1; y >= ymax; --y)
                        {
                            if (bitmap.GetPixel(x, y).ToArgb() != backgroundColor.Value.ToArgb())
                            {
                                ymax = Math.Max(ymax, y);
                            }
                        }
                    }

                    
                    //get codec based on GUID
                    var codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.FormatID == f.Guid);

                    //bitmap.Save(destination, codec, parameters);
                    // cropping result
                    bitmap.Clone(
                        new System.Drawing.Rectangle(xmin, ymin, xmax - xmin, ymax - ymin),
                        bitmap.PixelFormat
                    ).Save(destination, codec, parameters);

                }
            }
        }

        /// <summary>
        /// Save a metafile buffer (from inline shape range) to disk
        /// </summary>
        /// <param name="emfBuffer"></param>
        /// <param name="destinationFilePath"></param>
        public static void convertEmfBufferToPng(byte[] emfBuffer, string destinationFilePath)
        {
            using (var source = new MemoryStream(emfBuffer))
            {
                using (var destination = File.OpenWrite(destinationFilePath))
                {
                    SaveMetaFile(source, destination, 4, Color.White, ImageFormat.Png);
                }
            }
        }

        protected MSword.Application currentInstance;
        public DocumentPreprocessor(MSword.Application WordInstance) {
            currentInstance = WordInstance;
        }

        public ConversionStatus CreateWorkingCopy(ref object preprocessedObject, ref DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            MSword.Document currentDoc = (MSword.Document)preprocessedObject;
            object tmpFileName = document.CopyPath;

            if (File.Exists((string)tmpFileName)) {
                File.Delete((string)tmpFileName);
            }
            // Create an empty document to start the copy
            MSword.Document copy = currentInstance.Documents.Add(
                Visible: false
            );
            try {
                
                // Copy styles
                foreach (MSword.Style style in currentDoc.Styles) {
                    try {
                        currentInstance.OrganizerCopy(currentDoc.FullName, copy.FullName, style.NameLocal, MSword.WdOrganizerObject.wdOrganizerObjectStyles);
                    }
                    catch (Exception ex) {
                        AddinLogger.Warning("Non-critical exception raised while copying style " + style.NameLocal, ex);
                    }

                }

                // Copy properties
                DocumentProperties currentDocProps = (DocumentProperties)currentDoc.BuiltInDocumentProperties;
                DocumentProperties copyProps = (DocumentProperties)copy.BuiltInDocumentProperties;
                foreach (var key in Enum.GetValues(typeof(MSword.WdBuiltInProperty))) {
                    try {
                        copyProps[key].Value = currentDocProps[key].Value;
                    }
                    catch (Exception ex) {
                        // exception trigger on undefined property
                        AddinLogger.Warning("Non-critical exception raised copying built-in property " + key.ToString(), ex);
                    }

                }
                // Copy custom properties
                DocumentProperties currentDocCustomProps = (DocumentProperties)currentDoc.CustomDocumentProperties;
                DocumentProperties copyCustomProps = (DocumentProperties)copy.CustomDocumentProperties;
                foreach (DocumentProperty prop in currentDocCustomProps) {
                    try {
                        copyCustomProps.Add(prop.Name, prop.LinkToContent, prop.Type, prop.Value, prop.LinkSource);
                    }
                    catch (Exception ex) {
                        // exception trigger on undefined property
                        AddinLogger.Warning("Non-critical exception raised copying custom property " + prop.Name, ex);
                    }
                }

                // Copy content
                currentDoc.Content.Copy();
                copy.Content.Paste();
                Clipboard.Clear();

                // Save new document on disk
                copy.SaveAs2(
                    FileName: tmpFileName,
                    FileFormat: MSword.WdSaveFormat.wdFormatXMLDocument,
                    AddToRecentFiles: false
                );
                // use it as proprecessed document
                preprocessedObject = copy;
            } catch {
                // On error, close copy and forward exception
                copy.Close(SaveChanges: false);
                throw;
            }

            

            return ConversionStatus.CreatedWorkingCopy;

        }

        public ConversionStatus endPreprocessing(ref object preprocessedObject, IConversionEventsHandler eventsHandler = null) {
            MSword.Document preprocessingDocument = (MSword.Document)preprocessedObject;
            preprocessingDocument.Close(
                SaveChanges: MSword.WdSaveOptions.wdDoNotSaveChanges,
                OriginalFormat: MSword.WdOriginalFormat.wdOriginalDocumentFormat
            );
            return ConversionStatus.PreprocessingSucceeded;
        }

        public ConversionStatus ProcessEquations(ref object preprocessedObject, ref DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            Int16 showMsg = 0;
            MSword.Range rng;
            String storyName = "";
            int iNumShapesViewed = 0;
            MSword.Document currentDoc = (MSword.Document)preprocessedObject;

            foreach (MSword.Range tmprng in currentDoc.StoryRanges) {
                List<string> listmathML = new List<string>();
                rng = tmprng;
                storyName = rng.StoryType.ToString();
                while (rng != null) {
                    storyName = rng.StoryType.ToString();
                    MSword.InlineShapes shapes = rng.InlineShapes;
                    if (shapes != null && shapes.Count > 0) {
                        int iCount = 1;
                        int iNumShapes = 0;
                        Microsoft.Office.Interop.Word.InlineShape shape;
                        iNumShapes = shapes.Count;
                        //iCount is the LCV and the shapes accessor is 1 based, more that likely from VBA.

                        while (iCount <= iNumShapes) {
                            if (shapes[iCount].Type.ToString() == "wdInlineShapeEmbeddedOLEObject") {
                                if (shapes[iCount].OLEFormat.ProgID == "Equation.DSMT4") {
                                    shape = shapes[iCount];

                                    if (shape != null && shape.OLEFormat != null) {
                                        bool bRetVal = false;
                                        string strProgID;
                                        Guid autoConvert;
                                        strProgID = shape.OLEFormat.ProgID;
                                        bRetVal = GetFinalCLSID(ref strProgID, out autoConvert);

                                        // if we are successful with the conversion of the CLSID we now need to query
                                        //  the application to see if it can actually do the work
                                        if (bRetVal == true) {
                                            bool bInsertable = false;
                                            bool bNotInsertable = false;

                                            bInsertable = IsCLSIDInsertable(ref autoConvert);
                                            bNotInsertable = IsCLSIDNotInsertable(ref autoConvert);

                                            //Make sure that the server of interest is insertable and not-insertable
                                            if (bInsertable && bNotInsertable) {
                                                bool bServerExists = false;
                                                string strPathToExe;
                                                bServerExists = DoesServerExist(out strPathToExe, ref autoConvert);

                                                //if the server exists then see if MathML can be retrieved for the shape
                                                if (bServerExists) {
                                                    bool bMathML = false;
                                                    string strVerb;
                                                    int indexForVerb = -100;

                                                    strVerb = "RunForConversion";

                                                    bMathML = DoesServerSupportMathML(ref autoConvert, ref strVerb, out indexForVerb);
                                                    if (bMathML) {
                                                        storeMathMLEquation(ref shape, indexForVerb, listmathML);
                                                    }
                                                }
                                            } else {
                                                if (bInsertable != bNotInsertable) {
                                                    showMsg = 1;
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            //Increment the LCV and the number of shapes that iterated over.
                            iCount++;
                            iNumShapesViewed++;
                        }
                    }
                    rng = rng.NextStoryRange;
                }
                document.MathMLMap[storyName] = listmathML;
            }
            if (showMsg == 1) {
                string message =
                    "In order to convert MathType or Microsoft Equation Editor equations to DAISY,MathType 6.5 or later must be installed. See www.dessci.com/saveasdaisy for further information.Currently all the equations will be converted as Images";
                eventsHandler.OnStop(message, "Warning");
            }
            return ConversionStatus.ProcessedMathML;
        }

        public ConversionStatus ProcessShapes(ref object preprocessedObject, ref DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            MSword.Document currentDoc = (MSword.Document)preprocessedObject;
            //List<string> objectShapes = new List<string>();
            //List<string> imageIds = new List<string>();
            //List<string> inlineShapes = new List<string>();
            //List<string> inlineShapeIds = new List<string>();
            try {
                Exception threadEx = null;
                Thread staThread = new Thread(
                    delegate () {
                        try {
                            object missing = Type.Missing;
                            string outputPath = ConverterHelper.AppDataSaveAsDAISYDirectory;

                            List<string> warnings = new List<string>();
                            String fileName = currentDoc.Name.ToString().Replace(" ", "_");
                            MSword.Application WordInstance = currentDoc.Application;
                            WordInstance.Activate();

                            System.Diagnostics.Process objProcess = System.Diagnostics.Process.GetCurrentProcess();

                            foreach (MSword.Shape shape in currentDoc.Shapes) {
                                string name = shape.Name.ToString();
                                string type = shape.Type.ToString();
                                if (!shape.Name.Contains("Text Box")) {
                                    shape.Select(ref missing);
                                    string bookmark = "Shape_" + shape.ID.ToString();
                                    string shapeOutputPath = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(fileName) + "-Shape" + shape.ID.ToString() + ".png");
                                    WordInstance.Selection.CopyAsPicture();
                                    try {
                                        
                                        System.Drawing.Image image = ClipboardEx.GetEMF(objProcess.MainWindowHandle);
                                        byte[] Ret;
                                        MemoryStream ms = new MemoryStream();
                                        image.Save(ms, ImageFormat.Png);
                                        Ret = ms.ToArray();
                                        FileStream fs = new FileStream(shapeOutputPath, FileMode.Create, FileAccess.Write);

                                        fs.Write(Ret, 0, Ret.Length);
                                        fs.Flush();
                                        fs.Dispose();

                                        eventsHandler?.onFeedbackMessageReceived(this, new DaisyEventArgs(
                                            "Exported shape " + shapeOutputPath
                                        ));
                                        //objectShapes.Add(pathShape);
                                        //imageIds.Add(item.ID.ToString());
                                    } catch (ClipboardDataException cde) {
                                        warnings.Add("- Shape " + shape.ID.ToString() + ": " + cde.Message);
                                    } catch (Exception e) {
                                        throw e;
                                    } finally {
                                        Clipboard.Clear();
                                    }
                                }
                            }
                            MSword.Range rng;
                            foreach (MSword.Range tmprng in currentDoc.StoryRanges) {
                                rng = tmprng;
                                while (rng != null) {
                                    foreach (MSword.InlineShape item in rng.InlineShapes) {
                                        string type = item.Type.ToString();
                                        if ((item.Type.ToString() != "wdInlineShapeEmbeddedOLEObject") && ((item.Type.ToString() != "wdInlineShapePicture"))) {
                                            MSword.Shape shape = item.ConvertToShape();
                                            string bookmark = "Shape_" + shape.ID.ToString();
                                            string shapeOutputPath = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(fileName) + "-Shape" + shape.ID.ToString() + ".png");
                                            //object range = item.Range;
                                            //item.Range.Bookmarks.Add(bookmark, ref range);
                                            //item.Select();
                                            //item.Range.CopyAsPicture();
                                            try {
                                                byte[] buffer = (byte[])item.Range.EnhMetaFileBits;
                                                convertEmfBufferToPng(buffer, shapeOutputPath);
                                                eventsHandler?.onFeedbackMessageReceived(this, new DaisyEventArgs(
                                                    "Exported inlined shape " + shapeOutputPath
                                                ));
                                                //objectShapes.Add(shapeOutputPath);
                                                //imageIds.Add(shape.ID.ToString());
                                                //inlineShapes.Add(shapeOutputPath);
                                            } catch (ClipboardDataException cde) {
                                                warnings.Add("- InlineShape " + shape.ID.ToString() + " with AltText \"" + item.AlternativeText.ToString() + "\": " + cde.Message);
                                            } catch (Exception e) {
                                                throw e;
                                            } finally {
                                                Clipboard.Clear();

                                            }
                                        }
                                    }
                                    rng = rng.NextStoryRange;
                                }
                            }
                            if (warnings.Count > 0) {
                                string warningMessage = "Some shapes could not be exported from the document " + currentDoc.Name;
                                foreach (string warning in warnings) {
                                    warningMessage += "\r\n" + warning;
                                }
                                throw new Exception(warningMessage);
                            }
                        } catch (Exception ex) {
                            threadEx = ex;
                        }
                    });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join();
                if (threadEx != null) {
                    throw threadEx;
                }
                // not used
                //document.ObjectShapes = objectShapes;
                //document.InlineShapes = inlineShapes;
                //document.ImageIds = imageIds;
                //document.InlineIds = inlineShapeIds;
            } catch (Exception e) {
                eventsHandler?.OnError("An error occured while preprocessing shapes and may prevent the rest of the conversion to success:" +
                    "\r\n- " + e.Message +
                    "\r\n" + e.StackTrace);
            }
            return ConversionStatus.ProcessedShapes;
        }

        public object startPreprocessing(DocumentParameters document, IConversionEventsHandler eventsHandler = null) {
            // reset the focus on the document (or open it as visible) in the word app if it is not a subdoc
            return currentInstance.Documents.Open(
                FileName: document.InputPath,
                AddToRecentFiles: false,
                Visible: document.ResourceId == null
            );
        }

        public ConversionStatus ValidateName(ref object preprocessedObject, StringValidator authorizedNamePattern, IConversionEventsHandler eventsHandler = null) {
            MSword.Document currentDoc = (MSword.Document)preprocessedObject;
            MSword.Application WordInstance = currentDoc.Application;
            bool nameIsValid = false;
            do {
                bool docIsRenamed = false;
                if (!authorizedNamePattern.AuthorisationPattern.IsMatch(currentDoc.Name)) { // check only name (i assume it may still lead to problem if path has commas)

                    DialogResult? userAnswer = eventsHandler?.documentMustBeRenamed(authorizedNamePattern);
                    if (userAnswer.HasValue && userAnswer.Value == DialogResult.Yes) {
                        docIsRenamed = eventsHandler.userIsRenamingDocument(ref preprocessedObject);
                        if(!docIsRenamed) return ConversionStatus.Canceled;
                    } else if (userAnswer.HasValue && userAnswer.Value == DialogResult.Cancel) {
                        return ConversionStatus.Canceled;// PreprocessingData.Canceled("User canceled a renaming request for an invalid docx filename");
                    }
                    // else the sanitize path in the DaisyAddinLib will replace commas by underscore.
                    // this could be cleaner and done here instead of later to centralize file name related operations
                    // Other illegal characters regarding the conversion to DAISY book are replaced by underscore by the pipeline itself
                    // While image names seems to be sanitized in other process
                }
                nameIsValid = !docIsRenamed;
            } while (!nameIsValid);
            return ConversionStatus.ValidatedName;
        }

        #region Imports from dll and COMs
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern IntPtr GlobalLock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern bool GlobalUnlock(HandleRef handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int GlobalSize(HandleRef handle);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, ExactSpelling = true, SetLastError = true)]
        private static extern int OleGetAutoConvert(ref Guid oCurrentCLSID, out Guid pConvertedClsid);

        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool IsEqualGUID(ref Guid rclsid1, ref Guid rclsid);
        [ComImport]
        [Guid("00000112-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleObject {
            void DoNotCall_1(object x);
            void DoNotCall_2(ref object x);
            void SetHostNames(object szContainerApp, object szContainerObj);
            void Close(uint dwSaveOption);
        };

        public enum OLECLOSE {
            OLECLOSE_SAVEIFDIRTY = 0,
            OLECLOSE_NOSAVE = 1,
            OLECLOSE_PROMPTSAVE = 2,
        }
        #endregion

        #region MathType parsing requirements
        static private bool GetFinalCLSID(ref string ProgID, out Guid finalCLSID) {
            bool bRetVal = false;
            Guid oGuid;
            int iCOMRetVal = 0;

            iCOMRetVal = CLSIDFromProgID(ProgID, out oGuid);

            if (iCOMRetVal == 0) { // S_OK => Prog is associated to a Class
                FindFinalCLSID(ref oGuid, out finalCLSID);
                bRetVal = true;
            } else {
                finalCLSID = oGuid;
            }

            return bRetVal;
        }

        /// <summary>
        /// If a class belong to an "auto-conversion" chain of classes in the registry,
        /// get the final class of the chain.
        /// </summary>
        /// <param name="oldCLSID">original class ID</param>
        /// <param name="newCLSID">CLSID of the class that is at the end of the auto-conversion chain</param>
        static private void FindFinalCLSID(ref Guid oldCLSID, out Guid newCLSID) {
            int iCOMRetVal = 0;
            // Check if the class pointed by oldCLSID is set to be automatically to an other CLSID
            iCOMRetVal = OleGetAutoConvert(ref oldCLSID, out newCLSID);
            if (iCOMRetVal == 0) { // S_OK : no error during the call
                // Check if the new CLSID is not the old one (meaning we found the final CLSID)
                bool bGuidTheSame = false;
                try {
                    bGuidTheSame = IsEqualGUID(ref oldCLSID, ref newCLSID);
                } catch (COMException eCOM) {
                    MessageBox.Show(eCOM.Message);
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                }
                // Not the same ? keep going with the search
                if (bGuidTheSame == false) {
                    oldCLSID = newCLSID;
                    FindFinalCLSID(ref oldCLSID, out newCLSID);
                }
            } else {
                //There was some error in the auto conversion.
                // See if this guid will do the conversion.
                newCLSID = oldCLSID;
            }
        }

        static private bool IsCLSIDInsertable(ref Guid oGuid) {
            bool bInsertable = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;

            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"Insertable";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null)
                bInsertable = true;

            return bInsertable;
        }

        static private bool IsCLSIDNotInsertable(ref Guid oGuid) {
            bool bNotInsertable = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"NotInsertable";

            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            //The not-insertable key is not present.
            if (regkey == null)
                bNotInsertable = true;

            return bNotInsertable;
        }

        static private bool DoesServerExist(out string strPathToExe, ref Guid oGuid) {
            bool bServerExists = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;/* new Microsoft.Win32 Registry Key */
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + @"{" + oGuid.ToString() + @"}" + @"\" + @"LocalServer32";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetValueNames();
                strPathToExe = "";
                try {
                    strPathToExe = (string)regkey.GetValue(valnames[0]);
                } catch (Exception e) {
                }

                if (strPathToExe.Length > 0) {
                    //Now check if this is a good path
                    if (File.Exists(strPathToExe))
                        bServerExists = true;
                }

            } else {

                strPathToExe = null;

            }

            return bServerExists;
        }

        static private bool DoesServerSupportMathML(ref Guid oGuid, ref string strVerb, out int indexForVerb) {
            bool bIsMathMLSupported = false;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\DataFormats\GetSet";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                string[] valnames = regkey.GetSubKeyNames();
                int x = 0;
                while (x < regkey.SubKeyCount) {
                    RegistryKey subKey;
                    if (regkey.SubKeyCount > 0) {
                        subKey = regkey.OpenSubKey(valnames[x]);
                        if (subKey != null) {
                            string[] dataFormats = subKey.GetValueNames();
                            int y = 0;
                            while (y < subKey.ValueCount) {
                                string strValue = (string)subKey.GetValue(dataFormats[y]);

                                //This will accept both MathML and MathML Presentation.
                                if (strValue.Contains("MathML")) {
                                    bIsMathMLSupported = true;
                                    break;
                                }
                                y++;
                            }
                        }
                    }

                    if (bIsMathMLSupported)
                        break;
                    x++;
                }
            }

            //Now lets check to see if the appropriate verb is supported
            if (bIsMathMLSupported) {
                //The return value for a verb not found will be 1000
                //
                indexForVerb = GetVerbIndex(strVerb, ref oGuid);

                if (indexForVerb == 1000) {
                    bIsMathMLSupported = false;
                }
            } else {
                //We do not have an appropriate verb to start the server
                indexForVerb = -100;  //There is a predefined range for 
            }

            return bIsMathMLSupported;
        }

        static private int GetVerbIndex(string strVerbToFind, ref Guid oGuid) {
            int indexForVerb = 1000;
            //Check for the existance of the insertable key
            RegistryKey regkey;
            string strRegLocation;
            strRegLocation = @"Software\Classes\CLSID\" + "{" + oGuid.ToString() + "}" + @"\Verb";
            regkey = Registry.LocalMachine.OpenSubKey(strRegLocation);

            if (regkey != null) {
                //Lets make sure that we have some values before preceeding.
                if (regkey.SubKeyCount > 0) {
                    int x = 0;
                    int iCount = 0;

                    string[] valnames = regkey.GetSubKeyNames();

                    while (x < regkey.SubKeyCount) {
                        RegistryKey subKey;
                        if (regkey.SubKeyCount > 0) {
                            subKey = regkey.OpenSubKey(valnames[x]);
                            if (subKey != null) {
                                int y = 0;
                                string[] verbs = subKey.GetValueNames();
                                iCount = subKey.ValueCount;
                                string verb;

                                //Search all of the verbs for requested string.
                                while (y < iCount) {
                                    verb = (string)subKey.GetValue(verbs[y]);
                                    if (verb.Contains(strVerbToFind) == true) {
                                        string numVerb;
                                        numVerb = valnames[x].ToString();
                                        indexForVerb = int.Parse(numVerb);
                                        break;
                                    }
                                    y++;
                                }
                            }
                        }

                        //If the verb is not 1000 then break out of the verb
                        if (indexForVerb != 1000)
                            break;

                        x++;
                    }
                }
            }


            return indexForVerb;
        }

        static private void storeMathMLEquation(
            ref Microsoft.Office.Interop.Word.InlineShape shape,
            int indexForVerb,
            List<string> listmathML
        ) {
            IConnectDataObject mDataObject;
            if (shape != null) {
                object dataObject = null;
                object objVerb;

                objVerb = indexForVerb;

                //Start MathType, and get the dataobject that is connected to the server.    
                shape.OLEFormat.DoVerb(ref objVerb);

                try {
                    dataObject = shape.OLEFormat.Object;
                } catch (Exception e) {
                    //we have an issue with trying to get the verb,
                    //  There will be a attempt at another way to start the application.
                    MessageBox.Show(e.Message);
                }

                IOleObject oleObject = null;

                //This is a C# version of a QueryInterface
                if (dataObject != null) {
                    mDataObject = dataObject as IConnectDataObject;
                    oleObject = dataObject as IOleObject;
                } else {
                    //There was an issue with the addin trying to start with the verb we
                    // knew.  A backup is to call the with the primary verb and start the 
                    //  application normally.
                    objVerb = MSword.WdOLEVerb.wdOLEVerbPrimary;
                    shape.OLEFormat.DoVerb(ref objVerb);

                    dataObject = shape.OLEFormat.Object;
                    mDataObject = dataObject as IConnectDataObject;
                    oleObject = dataObject as IOleObject;
                }
                //Create instances of FORMATETC and STGMEDIUM for use with IDataObject
                ConnectFORMATETC oFormatEtc = new ConnectFORMATETC();
                ConnectSTGMEDIUM oStgMedium = new ConnectSTGMEDIUM();
                DataFormats.Format oFormat;



                //Find within the clipboard system the registered clipboard format for MathML
                oFormat = DataFormats.GetFormat("MathML");

                if (mDataObject != null) {
                    int iRetVal = 0;

                    //Initialize a FORMATETC structure to get the requested data
                    oFormatEtc.cfFormat = (Int16)oFormat.Id;
                    oFormatEtc.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
                    oFormatEtc.lindex = -1;
                    oFormatEtc.ptd = (IntPtr)0;
                    oFormatEtc.tymed = TYMED.TYMED_HGLOBAL;

                    iRetVal = mDataObject.QueryGetData(ref oFormatEtc);
                    //iRetVal will be zero if the MathML type is contained within the server.
                    if (iRetVal == 0) {
                        oStgMedium.tymed = TYMED.TYMED_NULL;
                    }

                    try {
                        mDataObject.GetData(ref oFormatEtc, out oStgMedium);
                    } catch (System.Runtime.InteropServices.COMException e) {
                        System.Windows.Forms.MessageBox.Show(e.ToString());
                        throw;
                    }

                    // Because we explicitly requested a MathML, we know that it is TYMED_HGLOBAL
                    // lets deal with the memory here.
                    if (oStgMedium.tymed == TYMED.TYMED_HGLOBAL &&
                        oStgMedium.unionmember != null) {
                        WriteOutMathMLFromStgMedium(ref oStgMedium, listmathML);

                        if (oleObject != null) {
                            uint close = (uint)OLECLOSE.OLECLOSE_NOSAVE;
                            // uint close = (uint)Microsoft.VisualStudio.OLE.Interop.OLECLOSE.OLECLOSE_NOSAVE;
                            oleObject.Close(close);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Store the mathML formated equation in a provided mathMl string array
        /// </summary>
        /// <param name="oStgMedium">memory handler of the equation</param>
        /// <param name="listmathML">String array to store the mathml in</param>
        static private void WriteOutMathMLFromStgMedium(
            ref ConnectSTGMEDIUM oStgMedium,
            List<string> listmathML
        ) {
            IntPtr ptr;
            byte[] rawArray = null;


            //Verify that our data contained within the STGMEDIUM is non-null
            if (oStgMedium.unionmember != null) {
                //Get the pointer to the data that is contained
                //  within the STGMEDIUM
                ptr = oStgMedium.unionmember;

                //The pointer now becomes a Handle reference.
                HandleRef handleRef = new HandleRef(null, ptr);

                try {
                    //Lock in the handle to get the pointer to the data
                    IntPtr ptr1 = GlobalLock(handleRef);

                    //Get the size of the memory block
                    int length = GlobalSize(handleRef);

                    //New an array of bytes and Marshal the data across.
                    rawArray = new byte[length];
                    Marshal.Copy(ptr1, rawArray, 0, length);

                    // I will now display the text.  Create a string from the rawArray
                    string str = Encoding.ASCII.GetString(rawArray);
                    str = str.Substring(str.IndexOf("<mml:math"), str.IndexOf("</mml:math>") - str.IndexOf("<mml:math"));
                    str = str + "</mml:math>";
                    str = str.Replace("xmlns:mml='http://www.w3.org/1998/Math/MathML'", "");

                    listmathML.Add(str);
                } catch (Exception exp) {
                    System.Diagnostics.Debug.WriteLine("MathMLimport from MathType threw an exception: " + Environment.NewLine + exp.ToString());
                } finally {
                    //This gets called regardless within a try catch.
                    //  It is a good place to clean up like this.
                    GlobalUnlock(handleRef);
                }
            }
        }
        #endregion
    }
}
