using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Text;
using System.Web;
using System.Xml.Xsl;
using System.Drawing;
using System.Xml.XPath;
using System.Collections;
using System.IO.Packaging;
using System.ComponentModel;
using System.IO.Compression;
using System.Drawing.Imaging;
using System.Collections.Generic;
using Daisy.DaisyConverter.DaisyConverterLib.PageStylesValidation;

namespace Daisy.DaisyConverter.DaisyConverterLib {
    /// <summary>
    /// Document-related functions wrapper class and global variable, to be called in XSLT
    /// </summary>
    class DaisyClass {

        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        const string footRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
        const string endRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
        const string appNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        const string appRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        const string numberRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        private String inputFilename, outputFilename, inputName, output_Pipeline;
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string CustomRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        const string customPropRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string emptyListCounter = "A";
        const String headingZeroLvl = "0";
        const String headingOneLvl = "1";
        const String headingSixLvl = "6";
        const String version2010 = "14.0", version2007 = "12.0", version2003 = "11.0", versionXP = "10.0";
        Package pack;

        /// <summary>
        /// stack of document levels value
        /// </summary>
        Stack<String> stackList = new Stack<String>();

        /// <summary>
        /// stack of levels value for lists
        /// </summary>
        Stack<String> lstackList = new Stack<String>();

        /// <summary>
        /// Stack of abbrevations (and acronyms)
        /// </summary>
        Stack<String> abbrstackList = new Stack<String>();

        /// <summary>
        /// Stack of lists headings
        /// </summary>
        Stack<String> listHeadingstackList = new Stack<String>();

        /// <summary>
        /// 
        /// </summary>
        Stack<String> abbrparastackList = new Stack<String>();

        /// <summary>
        /// 
        /// </summary>
        Stack<String> abbrheadstackList = new Stack<String>();

        /// <summary>
        /// 
        /// </summary>
        Stack<String> masteSubstackList = new Stack<String>();

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrHyperlink = new ArrayList();

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrListNote = new ArrayList();

        ArrayList arrList = new ArrayList(); // not used

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrListLang = new ArrayList();

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrCaptionProdnote = new ArrayList();
        String strImage = "", strImageExt = "";
        String sectionPagetype = "";
        String getAuthor = "", getTitle = "", getYear = "";
        String storeHyperId = "";
        String caption, message = "";
        String str, final, strImgName;
        Int16 AbbrAcrflag = 0;
        Int16 listflag = 0;
        Int16 listHeadingFlag = 0;
        Int16 imgId = 0;
        Int16 pageNum = 0;
        Int16 flagCheck = 0;
        Int16 length;
        int setHyperLinkFlag = 0, listMasterSubFlag = 0;
        int checkSectionFront = 0;
        int chekTocOccur = 0;

        /// <summary>
        /// Number of pages found before the toc
        /// </summary>
        int pageToc = 1;

        int checkToc = 0;

        /// <summary>
        /// 
        /// </summary>
        int checkSection = 0;

        /// <summary>
        /// 
        /// </summary>
        int checkSectionBody = 0;

        /// <summary>
        /// 
        /// </summary>
        int sectionCounter = 0;

        /// <summary>
        /// 
        /// </summary>
        int noteFlag = 0;

        /// <summary>
        /// 
        /// </summary>
        int incrementPage = 0;

        int rowspan = 0, setToc = 0, set_tabToc = 0, set_Toc = 0, bdoflag = 0, rtlFlag = 0;

        int captionFlag = 0, hyperlinkFlag = 0, testRun = 0;

        int set = 0, setbookmark = 0, checkCverpage = 0;

        int pageId = 0, flagcounter = 0;

        int sectionpageStart = 0;

        int codeFlag = 0;

        int conPageBreak = 0;

        int flagRowspan = 0;

        int linenumflag = 0;

        int tmpcount = 0;
        String prevHeadLvl = "";
        ArrayList startItem = new ArrayList(9);
        ArrayList prevHeadId = new ArrayList(2);
        Hashtable startHeadingItem = new Hashtable();
        ArrayList headingsNumId = new ArrayList();
        ArrayList bulletChar = new ArrayList();
        ArrayList headNumId = new ArrayList();
        ArrayList absFormat = new ArrayList();
        ArrayList OverideNumList = new ArrayList();
        String prevNumId = String.Empty;
        String prevHeadNumId = ""; String baseNumId = ""; String baseAbsId = "";
        Hashtable hTable = new Hashtable();
        Hashtable listCounters = new Hashtable();
        Hashtable headingCounters = new Hashtable();
        int objectId = 0;
        String absVal = "", numFormat = "", lvlText = "", lStartOverride = "", lStart = "", headingInfo = "";
        PackageRelationship relationship = null, imgRelationship = null, numberRelationship = null, customRelationship = null, footrelationship = null, endrelationship = null;
        Hashtable listMathMl;
        int footNoteFlag = 0, sidebarFlag = 0, mainFlag = 0, img_Flag = 0;

        /// <summary>
        /// DaisyClass constructor. 
        /// Initialize the global variables to be modified by function calls from xslt stylesheets.
        /// </summary>
        /// <param name="inputName">Input document name</param>
        /// <param name="input">Input document file (can be different from inputName, i.e. for temp file)</param>
        /// <param name="output">Output document file</param>
        /// <param name="listMathMl"></param>
        /// <param name="packInput">Input package</param>
        /// <param name="output_Pipeline">Output folder</param>
        public DaisyClass(String inputName, String input, String output, Hashtable listMathMl, Package packInput, string output_Pipeline) {
            this.inputName = inputName;
            inputFilename = input;
            outputFilename = output;
            pack = packInput;
            this.listMathMl = listMathMl;
            this.output_Pipeline = output_Pipeline;
            for (int i = 0; i < 9; i++) {
                startItem.Add("");
                absFormat.Add("");
            }

            for (int i = 0; i < 2; i++) {
                prevHeadId.Add("");
            }

            BulletChar();
        }

        /// <summary>
        /// - In the input document name (inputName), replace all spaces by underscore
        /// - If the output is set to the %APPDATA%/Daisy folder and if a png image ending by the "id" parameter exists in it, copy it to the "output_pipeline" folder
        /// - If the output is not the temporary folder and if a png image ending by the "id" parameter exists in %APPDATA%/SaveAsDAISY, move the image to the output folder
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string CheckShapeId(String id) {
            int flag = 0;
            inputName = inputName.Replace(" ", "_");
            if (string.Equals(outputFilename, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY\\")) {
                flag = 1;
                img_Flag = 1;
                string outputPath = output_Pipeline + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png";
                if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png")) {
                    File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png", output_Pipeline + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png", true);
                }
            }
            if (!string.Equals(Path.GetTempPath(), outputFilename + "\\") && flag == 0) {
                img_Flag = 1;
                string outputPath = outputFilename + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png";
                String tt = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png";
                if (File.Exists(tt)) {
                    string from = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png";
                    string to = outputFilename + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png";
                    if (from != to) {
                        File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png", outputFilename + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png", true);
                        File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + "SaveAsDAISY" + "\\" + Path.GetFileNameWithoutExtension(inputName) + "-" + id + ".png");
                    }
                }
            }
            id = Path.GetFileNameWithoutExtension(inputName) + "-" + id;
            id = id.Replace(" ", "_");
            return id;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="img"></param>
        /// <returns></returns>
        public string CheckImage(string img) {
            String tmp = outputFilename + "\\" + HttpUtility.UrlDecode(img);
            if (File.Exists(tmp)) {
                return "1";
            } else if (img_Flag == 1)
                return "0";
            else
                return "2";
        }


        /// <summary>
        /// Retrieve the next mathml for a story type.
        /// </summary>
        /// <param name="storyType">wdTextFrameStory, wdFootnotesStory or wdMainTextStory</param>
        /// <returns></returns>
        public String GetMathML(String storyType) {
            String strMathMl = "";
            if (listMathMl != null) {
                if (listMathMl.Count != 0) {
                    if (storyType == "wdTextFrameStory" && ((ArrayList)listMathMl[storyType]).Count != 0) {

                        if (((ArrayList)listMathMl[storyType])[sidebarFlag] != null) {
                            if (((ArrayList)listMathMl[storyType]).Count != 0) {
                                str = ((ArrayList)listMathMl[storyType])[sidebarFlag].ToString();
                                strMathMl = str;
                                sidebarFlag++;
                            }
                        }
                    } else if (storyType == "wdFootnotesStory" && ((ArrayList)listMathMl[storyType]).Count != 0) {
                        if (((ArrayList)listMathMl[storyType])[footNoteFlag] != null) {
                            if (((ArrayList)listMathMl[storyType]).Count != 0) {
                                str = ((ArrayList)listMathMl[storyType])[footNoteFlag].ToString();
                                strMathMl = str;
                                footNoteFlag++;
                            }
                        }
                    } else if (storyType == "wdMainTextStory" && ((ArrayList)listMathMl[storyType]).Count != 0) {
                        if (((ArrayList)listMathMl[storyType])[mainFlag] != null) {
                            if (((ArrayList)listMathMl[storyType]).Count != 0) {
                                str = ((ArrayList)listMathMl[storyType])[mainFlag].ToString();
                                strMathMl = str;
                                mainFlag++;
                            }
                        }
                    }
                }
            }
            return strMathMl;
        }

        /// <summary>
        /// Function Copying 2003/xp MathML Images from footnotes to output folder
        /// </summary>
        /// <param name="inNum"></param>
        /// <returns></returns>
        public String MathImageFootnote(string inNum) {
            try {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(footRelationshipType)) {
                    footrelationship = searchRelation;
                    break;
                }

                Uri footpartUri = PackUriHelper.ResolvePartUri(footrelationship.SourceUri, footrelationship.TargetUri);
                PackagePart footPartxml = pack.GetPart(footpartUri);

                imgRelationship = footPartxml.GetRelationship(inNum);
                Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
                PackagePart imgPartxml = pack.GetPart(imgpartUri);

                Image img = Bitmap.FromStream(imgPartxml.GetStream());

                Int16 length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('/').ToString());
                String strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                String str = outputFilename.Insert(outputFilename.Length, "\\");
                String final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);

                if (Path.GetExtension(final) != ".jpeg" && Path.GetExtension(final) != ".png") {
                    img.Save(final.Replace(Path.GetExtension(final), ".png"), ImageFormat.Png);
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1).Replace(Path.GetExtension(final),
                        ".png");

                } else if (Path.GetExtension(final) == ".jpeg") {
                    img.Save(final.Replace(Path.GetExtension(final), ".jpg"));
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1).Replace(Path.GetExtension(final), ".jpg");
                } else {
                    img.Save(final);
                }

                return UriEscape(Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
            } catch {
                return "translation.oox2Daisy.ImageContent";
            }
        }

        /// <summary>
        /// Function Copying 2003/xp MathML Images to output folder
        /// </summary>
        /// <param name="inNum"></param>
        /// <returns></returns>
        public String MathImage(string inNum) {
            try {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                imgRelationship = mainPartxml.GetRelationship(inNum);
                Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
                PackagePart imgPartxml = pack.GetPart(imgpartUri);

                Image img = Bitmap.FromStream(imgPartxml.GetStream());

                Int16 length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('/').ToString());
                String strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                String str = outputFilename.Insert(outputFilename.Length, "\\");
                String final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);

                if (Path.GetExtension(final) != ".jpeg" && Path.GetExtension(final) != ".png") {
                    img.Save(final.Replace(Path.GetExtension(final), ".png"), ImageFormat.Png);
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1).Replace(Path.GetExtension(final), ".png");
                } else {
                    img.Save(final);
                }

                return UriEscape(Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
            } catch {
                return "translation.oox2Daisy.ImageContent";
            }
        }

        /// <summary>
        /// Function to get the image file names for Office Word 2007
        /// </summary>
        /// <param name="inNum"></param>
        /// <param name="imageName"></param>
        /// <param name="val"></param>
        /// <returns></returns>
        public String Image(String inNum, string imageName, String val) {
            try {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                imgRelationship = mainPartxml.GetRelationship(inNum);
                Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
                PackagePart imgPartxml = pack.GetPart(imgpartUri);

                Image img = Bitmap.FromStream(imgPartxml.GetStream());
                string retimg2007Name;

                /*Checking if full filename(along with extn) of the image exists*/
                if (Path.GetExtension(imageName) != "") {
                    strImgName = imageName;
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
                    retimg2007Name = ImageExt(inNum, final, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
                    return UriEscape(retimg2007Name);
                }
                /*Checking if full filename(along with extn) of the image does not exist*/
                else if (Path.GetExtension(imageName) == "" && imageName != "") {
                    length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('.').ToString());
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                    string img2007Name = imageName + "." + strImgName;
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + img2007Name);
                    retimg2007Name = ImageExt(inNum, final, Path.GetFileNameWithoutExtension(inputName) + "-" + img2007Name);
                    return UriEscape(retimg2007Name);
                }
                /*Checking if entire filename of the image doesn't exist*/
                else {
                    length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('/').ToString());
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    strImgName = Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName;
                    final = str.Insert(str.Length, strImgName);
                    retimg2007Name = ImageExt(inNum, final, strImgName);
                    return UriEscape(retimg2007Name);
                }
            } catch {
                return "translation.oox2Daisy.ImageContent";
            }
        }

        /// <summary>
        /// Function to resample images in the document according to daisy settings
        /// </summary>
        /// <param name="inNum"></param>
        /// <param name="imageName"></param>
        /// <param name="resampleValue"></param>
        /// <returns></returns>
        public String ResampleImage(String inNum, string imageName, String resampleValue) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            imgRelationship = mainPartxml.GetRelationship(inNum);
            Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);

            PackagePart imgPartxml = pack.GetPart(imgpartUri);

            //Shaby : Implementation of Image Sampling
            string srcName = Path.GetFileNameWithoutExtension(imageName); ;
            string srcFormat = Path.GetExtension(imgpartUri.ToString()).Remove(0, 1);
            Image srcImg = Bitmap.FromStream(imgPartxml.GetStream());

            ImageProcessing sourceImage = new ImageProcessing(srcImg, srcFormat, srcName);
            ImageProcessing targetImage = new ImageProcessing();
            targetImage.Name = sourceImage.Name;
            targetImage.Format = sourceImage.Format;
            targetImage.Resolution = (float)Convert.ToDecimal(resampleValue);
            targetImage.Width = sourceImage.XResolutionRate * targetImage.Resolution;
            targetImage.Height = sourceImage.YResolutionRate * targetImage.Resolution;
            targetImage.SrcImage = sourceImage.SrcImage;
            try {
                return targetImage.SaveProcessedImage(outputFilename);
            } catch (Exception ex) {
                throw new Exception(ex.Message, ex.InnerException);
            }
        }


        /// <summary>
        /// Function to Copy image to the Output/destination folder
        /// </summary>
        /// <param name="inNum"></param>
        /// <param name="pathName"></param>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public String ImageExt(String inNum, String pathName, String imageName) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            imgRelationship = mainPartxml.GetRelationship(inNum);
            Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
            PackagePart imgPartxml = pack.GetPart(imgpartUri);

            Image img = Bitmap.FromStream(imgPartxml.GetStream());

            /*Checking if filename extension is not .jpeg, .jpg, .jpe,.png and converting it to filename with .png extn*/
            if ((Path.GetExtension(pathName) != ".jpeg") && (Path.GetExtension(pathName) != ".jpe") && (Path.GetExtension(pathName) != ".jpg") && (Path.GetExtension(pathName) != ".JPEG") && (Path.GetExtension(pathName) != ".JPE") && (Path.GetExtension(pathName) != ".JPG") && (Path.GetExtension(pathName) != ".png") && (Path.GetExtension(pathName) != ".PNG")) {

                img.Save(pathName.Replace(Path.GetExtension(pathName), ".png"), ImageFormat.Png);
                strImgName = imageName.Replace(Path.GetExtension(pathName), ".png");
                return strImgName;

            }
            /*Checking if filename extension is .jpeg, .jpg, .jpe and converting it to filename with .jpg extn*/
            else if ((Path.GetExtension(pathName) == ".jpeg") || (Path.GetExtension(pathName) == ".jpg") || (Path.GetExtension(pathName) == ".jpe") || (Path.GetExtension(pathName) == ".JPE") || (Path.GetExtension(pathName) == ".JPEG") || (Path.GetExtension(pathName) == ".JPG")) {
                //img.Save(pathName.Replace(Path.GetExtension(pathName), ".jpg"));

                Bitmap imgbitmap = new Bitmap(img);
                imgbitmap.Save(pathName.Replace(Path.GetExtension(pathName), ".jpg"), ImageFormat.Jpeg);

                strImgName = imageName.Replace(Path.GetExtension(pathName), ".jpg");
                return strImgName;
            }
            /*If filename extension is .png it is returned as it is*/
            else {
                img.Save(pathName);
                return imageName;
            }
        }

        /// <summary>
        /// Function to get the image file names for Office Word 2003 and Grouped Images
        /// </summary>
        /// <param name="inNum"></param>
        /// <param name="imagegroupName"></param>
        /// <returns></returns>
        public String Image(string inNum, string imagegroupName) {
            try {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                imgRelationship = mainPartxml.GetRelationship(inNum);
                Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
                PackagePart imgPartxml = pack.GetPart(imgpartUri);

                Image img = Bitmap.FromStream(imgPartxml.GetStream());
                string imggrpName, retimggrpName;

                /*Checking if full filename(along with extn) of the image exists*/
                if (Path.GetExtension(imagegroupName) != "") {
                    strImgName = imagegroupName;
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
                    retimggrpName = ImageExt(inNum, final, Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName);
                    return UriEscape(retimggrpName);
                }
                /*Checking if full filename(along with extn) of the image does not exist*/
                else if (Path.GetExtension(imagegroupName) == "" && imagegroupName != "") {
                    length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('.').ToString());
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                    imggrpName = imagegroupName + "." + strImgName;
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    final = str.Insert(str.Length, Path.GetFileNameWithoutExtension(inputName) + "-" + imggrpName);
                    retimggrpName = ImageExt(inNum, final, Path.GetFileNameWithoutExtension(inputName) + "-" + imggrpName);
                    return UriEscape(retimggrpName);
                }
                /*Checking if entire filename of the image doesn't exist*/
                else if (imagegroupName == "") {
                    length = Int16.Parse(imgPartxml.Uri.ToString().LastIndexOf('/').ToString());
                    strImgName = imgPartxml.Uri.ToString().Substring(length + 1);
                    str = outputFilename.Insert(outputFilename.Length, "\\");
                    imggrpName = Path.GetFileNameWithoutExtension(inputName) + "-" + strImgName;
                    final = str.Insert(str.Length, imggrpName);
                    retimggrpName = ImageExt(inNum, final, imggrpName);
                    return UriEscape(retimggrpName);
                } else {
                    return "0";
                }

            } catch {
                return "translation.oox2Daisy.ImageContent";
            }
        }


        /// <summary>
        /// Function to check Document is having external images or not
        /// </summary>
        /// <returns></returns>
        public String ExternalImage() {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships()) {
                relationship = searchRelation;
                if (relationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
                    if (relationship.TargetMode.ToString() == "External") {
                        strImageExt = "translation.oox2Daisy.ExternalImage";
                        break;
                    }
                }

                strImageExt = "no external image";
            }
            return strImageExt;
        }

        /// <summary>
        /// Function to Push current document level value to the general stack
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        public String PushLevel(String level) {
            
            if (stackList.Count > 0) {
                if (level.Length > 0 && Convert.ToInt16(level) > 6 && stackList.Count == 6) {
                    stackList.Push(headingSixLvl);
                } else if (level != stackList.Peek() + 1) {
                    level = Convert.ToString(Convert.ToInt16(stackList.Peek()) + 1);
                    stackList.Push(level);
                } else {
                    stackList.Push(level);
                }

            } else {
                if (level != headingOneLvl) {
                    level = headingOneLvl;
                    stackList.Push(level);
                } else {
                    stackList.Push(level);
                }
            }
            return level;
        }

        /// <summary>
        /// peek (no suppresion) the last level of the document from the stack
        /// </summary>
        /// <returns></returns>
        public String PeekLevel() {
            if (stackList.Count > 0) {
                return stackList.Peek();
            } else {
                return headingZeroLvl;
            }
        }

        /// <summary>
        /// pop (retrieve and delete) the last level of the document from the stack
        /// </summary>
        /// <returns></returns>
        public String PoPLevel() {
            return stackList.Pop();
        }

        /// <summary>
        /// Increment and return as a string the pageToc counter (number of pages before the toc)
        /// </summary>
        /// <returns></returns>
        public string PageForTOC() {
            pageToc++;
            return pageToc.ToString();
        }

        /// <summary>
        /// Function which returns whether TOC is there or not.
        /// (Also increment a checkToc counter)
        /// </summary>
        /// <returns></returns>
        public string CheckToc() {
            checkToc++;
            return checkToc.ToString();
        }

        /// <summary>
        /// Function which returns Upper Roman letter with respect to an integer
        /// </summary>
        /// <param name="counter"></param>
        /// <returns></returns>
        public string PageNumUpperRoman(string counter) {
            Int16[] values = new Int16[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] numerals = new string[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            StringBuilder result = new StringBuilder();

            Int16 check = Convert.ToInt16(counter);
            if (check == 0) {
                check = 1;
            }
            for (int i = 0; i < values.Length; i++) {
                // If the number being converted is less than the test value, append
                // the corresponding numeral or numeral pair to the resultant string
                while (check >= values[i]) {
                    check -= values[i];
                    result.Append(numerals[i]);
                }
            }
            return result.ToString(); ;
        }

        /// <summary>
        /// Function which returns Lower Roman letter with respect to an integer
        /// </summary>
        /// <param name="counter"></param>
        /// <returns></returns>        
        public string PageNumLowerRoman(string counter) {

            Int16[] values = new Int16[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] numerals = new string[] { "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i" };
            StringBuilder result = new StringBuilder();
            Int16 check = Convert.ToInt16(counter);
            if (check == 0) {
                check = 1;
            }
            for (int i = 0; i < values.Length; i++) {
                // If the number being converted is less than the test value, append
                // the corresponding numeral or numeral pair to the resultant string
                while (check >= values[i]) {
                    check -= values[i];
                    result.Append(numerals[i]);
                }
            }
            return result.ToString(); ;
        }


        /// <summary>
        /// Function which returns Lower Alphabet with respect to an integer
        /// </summary>
        /// <param name="counter"></param>
        /// <returns></returns>
        public string PageNumLowerAlphabet(string counter) {
            string[] numerals = new string[] { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
            string lowerAlpha;
            int check = Convert.ToInt16(counter);
            /*if counter value is greater than 26,then checking the difference and getting the proper alphabet*/
            if (check > 26) {
                check = check - 26;
                lowerAlpha = string.Concat(numerals[check - 1], numerals[check - 1]);
            } else if (check != 0) {
                lowerAlpha = numerals[check - 1];
            } else {
                lowerAlpha = numerals[0];
            }
            return lowerAlpha;
        }

        /// <summary>
        /// Function which returns Upper Alphabet with respect to an integer
        /// </summary>
        /// <param name="counter"></param>
        /// <returns></returns>
        public string PageNumUpperAlphabet(string counter) {
            string[] numerals = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            string upperAlpha;
            int check = Convert.ToInt16(counter);
            /*if counter value is greater than 26,then checking the difference and getting the proper alphabet*/
            if (check > 26) {
                check = check - 26;
                upperAlpha = string.Concat(numerals[check - 1], numerals[check - 1]);
            } else if (check != 0) {
                upperAlpha = numerals[check - 1];
            } else {
                upperAlpha = numerals[0];
            }
            return upperAlpha;

        }

        /// <summary>
        /// Function which returns the flag regarding PageNumbers
        /// </summary>
        /// <returns></returns>
        public string CheckForcounter() {
            flagcounter++;
            return flagcounter.ToString();
        }

        /// <summary>
        /// Function which returns the number of Sections in the document
        /// </summary>
        /// <param name="pageType"></param>
        /// <param name="pageStart"></param>
        /// <returns></returns>
        public string SectionCounter(string pageType, string pageStart) {
            if (string.IsNullOrEmpty(pageType)) {
                sectionPagetype = "decimal";
            } else {
                sectionPagetype = pageType;
            }
            if (string.IsNullOrEmpty(pageStart)) {
                pageNum = 1;
            } else {
                pageNum = Convert.ToInt16(pageStart);
            }

            sectionCounter++;
            return sectionCounter.ToString();
        }

        /// <summary>
        /// Function which returns the type of page format
        /// </summary>
        /// <returns></returns>
        public string GetPageFormat() {
            if (string.IsNullOrEmpty(sectionPagetype)) {
                return "decimal";
            } else {
                return sectionPagetype;
            }
        }

        /// <summary>
        /// Function which returns the Section value
        /// </summary>
        /// <returns></returns>
        public string GetSectionCounter() {
            return sectionCounter.ToString();
        }

        /// <summary>
        /// Function which returns the Page Number
        /// </summary>
        /// <returns></returns>
        public string GetPageNum() {
            return pageNum.ToString();
        }

        /// <summary>
        /// Function which returns the Section value for the first page
        /// </summary>
        /// <returns></returns>
        public string GetSectionPageStart() {
            sectionpageStart++;
            return sectionpageStart.ToString();
        }

        /// <summary>
        /// Function which initializes the Section value for page Start
        /// </summary>
        /// <returns></returns>
        public string InitalizeSectionPageStart() {
            sectionpageStart = 0;
            return sectionpageStart.ToString();
        }

        /// <summary>
        /// Function which initializes the Section value for page body
        /// </summary>
        /// <returns></returns>
        public string InitalizeSectionPageBody() {
            checkSectionBody = 0;
            return checkSectionBody.ToString();
        }

        /// <summary>
        /// Function which increments the TOC counter
        /// </summary>
        /// <returns></returns>
        public string CheckTocOccur() {
            chekTocOccur++;
            return chekTocOccur.ToString();
        }

        /// <summary>
        /// Function which increments the Section counter
        /// </summary>
        /// <returns></returns>
        public string CheckSection() {
            checkSection++;
            return checkSection.ToString();
        }

        /// <summary>
        /// Function which initializes the Section value
        /// </summary>
        /// <returns></returns>
        public string InitalizeCheckSection() {
            checkSection = 0;
            return checkSection.ToString();
        }

        /// <summary>
        /// Function which increments the Section counter for BODY
        /// </summary>
        /// <returns></returns>
        public string CheckSectionBody() {
            checkSectionBody++;
            return checkSectionBody.ToString();
        }

        /// <summary>
        /// Function which initializes the Section counter for BODY
        /// </summary>
        /// <returns></returns>
        public string InitalizeCheckSectionBody() {
            checkSectionBody = 0;
            return checkSectionBody.ToString();
        }

        /// <summary>
        /// Function which returns the page number
        /// </summary>
        /// <param name="start"></param>
        /// <returns></returns>
        public string PageNum(string start) {
            pageNum = Convert.ToInt16(start);
            return pageNum.ToString();

        }

        /// <summary>
        /// Function which increments the page number
        /// </summary>
        /// <returns></returns>
        public string IncrementPageNo() {
            pageNum++;
            return pageNum.ToString();
        }

        /// <summary>
        /// Function which increments page occurence counter
        /// </summary>
        /// <returns></returns>
        public string CheckPageOccurance() {
            flagCheck++;
            return flagCheck.ToString();
        }

        /// <summary>
        /// Function which returns page Id
        /// </summary>
        /// <returns></returns>
        public string GeneratePageId() {
            pageId++;
            return pageId.ToString();
        }

        /// <summary>
        /// Function which increments page occurence counter
        /// </summary>
        /// <returns></returns>
        public string IncrementPage() {
            incrementPage++;
            return incrementPage.ToString();
        }

        /// <summary>
        /// Function which returns page number
        /// </summary>
        /// <returns></returns>
        public string ReturnPageNum() {
            return incrementPage.ToString();
        }

        /// <summary>
        /// Function which increments Section counter in front matter
        /// </summary>
        /// <returns></returns>
        public string CheckSectionFront() {
            checkSectionFront++;
            return checkSectionFront.ToString();
        }

        /// <summary>
        /// Function which returns Section counter in front matter
        /// </summary>
        /// <returns></returns>
        public string GetSectionFront() {
            return checkSectionFront.ToString();
        }


        /// <summary>
        /// Function to add Footnote to an Array
        /// </summary>
        /// <param name="inNum"></param>
        /// <returns></returns>
        public String AddFootNote(string inNum) {
            arrListNote.Add(inNum);
            return inNum;
        }

        /// <summary>
        /// Function to get Value of a particular Footnote
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public String FootNoteId(int i) {
            if (arrListNote.Count > 0) {
                String s1 = arrListNote[i].ToString();
                arrListNote.RemoveAt(i);
                return s1;
            } else {
                return ("0");
            }

        }

        /// <summary>
        /// Function which returns the flag regarding Footnotes
        /// </summary>
        /// <returns></returns>
        public string NoteFlag() {
            noteFlag++;
            return noteFlag.ToString();
        }

        /// <summary>
        /// Function which initializes the flag regarding Footnotes
        /// </summary>
        /// <returns></returns>
        public string InitializeNoteFlag() {
            noteFlag = 0;
            return noteFlag.ToString();
        }

        /// <summary>
        /// Function to add Caption and Prodnote to an Array
        /// </summary>
        /// <returns></returns>
        public int AddCaptionsProdnotes() {
            tmpcount++;
            arrCaptionProdnote.Add(tmpcount);
            return tmpcount;
        }


        /// <summary>
        /// Function to get Caption and Prodnote from an Array
        /// </summary>
        /// <returns></returns>
        public int GetCaptionsProdnotes() {
            if (arrCaptionProdnote.Count > 0) {
                int cnt = Convert.ToInt32(arrCaptionProdnote[0]);
                arrCaptionProdnote.RemoveAt(0);
                return cnt;
            } else {
                return (0);
            }

        }


        /// <summary>
        /// Function to reset the count value used in the AddCaptionsProdnotes function
        /// </summary>
        public void ResetCaptionsProdnotes() {
            tmpcount = 0;
        }


        /// <summary>
        /// Function to Push a level value to the stack 
        /// (only if the level is superior to the last one in stack)
        /// Used in Common2.xsl:411
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public int ListPush(int i) {
            string j = i.ToString();
            if (lstackList.Count > 0) {
                if (i > Convert.ToInt16(lstackList.Peek())) {
                    j = lstackList.Peek();
                    lstackList.Push(i.ToString());
                }
            } else {
                lstackList.Push(i.ToString());
            }
            return Convert.ToInt16(j);
        }


        /// <summary>
        /// Function to Peek the top value of the Stack
        /// </summary>
        /// <returns></returns>
        public int ListPeekLevel() {
            if (lstackList.Count > 0) {
                return Convert.ToInt16(lstackList.Peek());
            } else {
                return (0);
            }
        }

        /// <summary>
        /// Function to PoP the top value of the Stack
        /// </summary>
        /// <returns></returns>
        public int ListPoPLevel() {
            if (lstackList.Count > 0)
                return Convert.ToInt16(lstackList.Pop());
            else
                return 0;
        }

        /// <summary>
        /// Function to generate a new ID for image
        /// (increment the id counter)
        /// </summary>
        /// <returns></returns>
        public int GenerateImageId() {
            return ++imgId;
        }

        /// <summary>
        /// Function to insert Caption for an Image
        /// </summary>
        /// <param name="captionImg"></param>
        /// <returns></returns>
        public String InsertCaption(String captionImg) {
            caption = captionImg;
            return caption;
        }

        /// <summary>
        /// Function to get Id of a particular Image
        /// </summary>
        /// <param name="imgRef"></param>
        /// <returns></returns>
        public String ImageGroupId(string imgRef) {

            strImage = String.Concat(strImage, imgRef, GenerateImageId(), " ");
            return strImage;
        }

        /// <summary>
        /// Function returns image index
        /// </summary>
        /// <returns></returns>
        public int storeImgid() {
            int indexImage = imgId;
            return indexImage;
        }


        /// <summary>
        /// Function returs caption
        /// </summary>
        /// <returns></returns>
        public String ReturnCaption() {
            if (string.IsNullOrEmpty(caption)) {
                return ("0");
            } else {
                return caption;
            }
        }

        /// <summary>
        /// Function to get ID of an image
        /// </summary>
        /// <returns></returns>
        public String ReturnImageGroupId() {
            return strImage;
        }

        /// <summary>
        /// Function to get Unique ID
        /// </summary>
        /// <returns></returns>
        public long GenerateId() {
            byte[] buffer = Guid.NewGuid().ToByteArray();
            return BitConverter.ToInt64(buffer, 0);
        }

        /// <summary>
        /// Function to returns flag value
        /// </summary>
        /// <returns></returns>
        public int Flag() {
            int flag = 0;
            return flag;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="checkLvl"></param>
        /// <returns></returns>
        public int GetCheckLvlInt(XPathNodeIterator checkLvl) {
            try {
                if (checkLvl.Count > 0) {
                    checkLvl.MoveNext();
                    return checkLvl.Current.ValueAsInt;
                } else {
                    return 0;
                }
            } catch {
                throw;
            }
        }

        /// <summary>
        /// Function which gives the difference between Current level and Peek level of Stack
        /// </summary>
        /// <param name="close"></param>
        /// <param name="peekLevel"></param>
        /// <returns></returns>
        public bool Difference(int close, int peekLevel) {
            if (peekLevel - close > 1)
                return true;

            return false;
        }

        /// <summary>
        /// Function which decrements the counter value
        /// </summary>
        /// <param name="rec"></param>
        /// <returns></returns>
        public int Decrement(int rec) {
            return rec - 1;
        }

        /// <summary>
        /// Function which increments the counter value
        /// </summary>
        /// <param name="rec"></param>
        public void Increment(int rec) {
            for (int i = 0; i <= rec; i++)
                ListPush(i);
        }

        /// <summary>
        /// Function returns the target string of an anchor
        /// </summary>
        /// <param name="inNum"></param>
        /// <returns></returns>
        public String Anchor(string inNum) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);
            imgRelationship = mainPartxml.GetRelationship(inNum);
            String uri = imgRelationship.TargetUri.ToString();
            return HttpUtility.UrlPathEncode(uri);
        }

        /// <summary>
        /// Function which concatenate messages for TOC
        /// </summary>
        /// <param name="str"></param>
        public void SetTOCMessage(string str) {
            message = message + " " + str;
        }

        /// <summary>
        /// Function which return message for TOC
        /// </summary>
        /// <returns></returns>
        public string GetTOCMessage() {
            return message;
        }

        /// <summary>
        /// Function which returns Null message for TOC
        /// </summary>
        public void NullMsg() {
            message = "";
        }

        /// <summary>
        /// Function which increments the counter for TOC
        /// </summary>
        /// <returns></returns>
        public string SetToc() {
            setToc++;
            return setToc.ToString();
        }

        /// <summary>
        /// Function which resets the counter value for TOC
        /// </summary>
        /// <returns></returns>
        public string GetToc() {
            setToc = 0;
            return setToc.ToString();
        }

        /// <summary>
        /// Function which increments the counter for TOC
        /// </summary>
        /// <returns></returns>
        public string Set_tabToc() {
            set_tabToc++;
            return set_tabToc.ToString();
        }

        /// <summary>
        /// Function which resets the counter value for TOC
        /// </summary>
        /// <returns></returns>
        public string Get_tabToc() {
            set_tabToc = 0;
            return set_tabToc.ToString();
        }


        /// <summary>
        /// Function which increments the counter for TOC
        /// </summary>
        /// <returns></returns>
        public string Set_Toc() {
            set_Toc++;
            return set_Toc.ToString();
        }

        /// <summary>
        /// Function which resets the counter value for TOC
        /// </summary>
        /// <returns></returns>
        public string Get_Toc() {
            set_Toc = 0;
            return set_Toc.ToString();
        }

        /// <summary>
        /// Function which sets the counter value for Continuous page break
        /// </summary>
        /// <returns></returns>
        public string SetConPageBreak() {
            conPageBreak++;
            return conPageBreak.ToString();
        }


        /// <summary>
        /// Function which Resetsets the counter value for Continuous page break
        /// </summary>
        /// <returns></returns>
        public string ResetSetConPageBreak() {
            conPageBreak = 0;
            return conPageBreak.ToString();
        }

        /// <summary>
        /// Functions to Implement Citation
        /// </summary>
        /// <returns></returns>
        public String Citation() {

            String indicator = " ";
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(CustomRelationshipType)) {
                customRelationship = searchRelation;
                Uri CustomUri = PackUriHelper.ResolvePartUri(customRelationship.SourceUri, customRelationship.TargetUri);
                PackagePart customPartxml = pack.GetPart(CustomUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(customPartxml.GetStream());
                if (doc.DocumentElement.NamespaceURI == "http://schemas.openxmlformats.org/officeDocument/2006/bibliography") {
                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("b", "http://schemas.openxmlformats.org/officeDocument/2006/bibliography");
                    XmlNode node = doc.SelectSingleNode("//b:Sources/@StyleName", nsManager);

                    if (node != null) {
                        indicator = node.InnerText;
                    } else
                        indicator = " ";
                }
            }
            return indicator;
        }


        /// <summary>
        /// Function returns details(title,author,year) of a citation
        /// </summary>
        /// <param name="citeId"></param>
        /// <returns></returns>
        public String CitationDetails(string citeId) {
            string temp = "";
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);
            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(CustomRelationshipType)) {
                customRelationship = searchRelation;
                Uri CustomUri = PackUriHelper.ResolvePartUri(customRelationship.SourceUri, customRelationship.TargetUri);
                PackagePart customPartxml = pack.GetPart(CustomUri);
                XmlDocument doc = new XmlDocument();
                doc.Load(customPartxml.GetStream());
                if (doc.DocumentElement.NamespaceURI == "http://schemas.openxmlformats.org/officeDocument/2006/bibliography") {
                    NameTable nt = new NameTable();
                    XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                    nsManager.AddNamespace("b", "http://schemas.openxmlformats.org/officeDocument/2006/bibliography");
                    XmlNodeList list = doc.SelectNodes("//b:Source/b:Tag", nsManager);
                    for (int i = 0; i < list.Count; i++) {

                        if (citeId.Contains(list.Item(i).InnerText)) {
                            XmlNode getTitle1 = list.Item(i).ParentNode.SelectSingleNode("b:Title", nsManager);
                            if (getTitle1 != null)
                                getTitle = getTitle1.InnerText;
                            XmlNode getYear1 = list.Item(i).ParentNode.SelectSingleNode("b:Year", nsManager);
                            if (getYear1 != null)
                                getYear = getYear1.InnerText;
                            XmlNodeList listAuthor = list.Item(i).ParentNode.SelectNodes("b:Author/b:Author/b:NameList//b:Person/b:Last", nsManager);
                            if (listAuthor != null) {
                                for (int j = 0; j < listAuthor.Count; j++) {
                                    temp = temp + " " + listAuthor.Item(j).InnerText;
                                }
                                getAuthor = temp;
                            }
                        }

                    }

                }
            }
            return " ";
        }

        /// <summary>
        /// Function returns Author name of a citation
        /// </summary>
        /// <returns></returns>
        public string GetAuthor() {
            return getAuthor;
        }

        /// <summary>
        /// Function returns Title of a citation
        /// </summary>
        /// <returns></returns>
        public string GetTitle() {
            return getTitle;
        }

        /// <summary>
        /// Function returns Year of the citation
        /// </summary>
        /// <returns></returns>
        public string GetYear() {
            return getYear;
        }

        /// <summary>
        /// Function to implement Language
        /// </summary>
        /// <param name="lang"></param>
        /// <returns></returns>
        public string AddLanguage(string lang) {
            int flag = 1;
            if (arrListLang.Count == 0) {
                arrListLang.Add(lang);
            } else if (arrListLang.Count > 0) {
                for (int count = 0; count <= arrListLang.Count - 1; count++) {
                    if (string.Compare(lang, arrListLang[count].ToString()) == 0) {
                        flag = 0;
                    }
                }

                if (flag == 1) {
                    arrListLang.Add(lang);
                }
            }
            return flag.ToString();

        }

        /// <summary>
        /// Functions for HyperLink
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string AddHyperlink(string name) {
            if (string.IsNullOrEmpty(name)) {
                name = "";
            } else {
                arrHyperlink.Add(name);
            }
            return name;
        }

        /// <summary>
        /// Function to check whether given name is hyperlink or not
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public string GetHyperlinkName(string name) {
            int flag = 0;
            if (string.IsNullOrEmpty(name) && arrHyperlink.Count == 0) {
                flag = 0;
            } else if (arrHyperlink.Count > 0) {
                for (int count = 0; count <= arrHyperlink.Count - 1; count++) {
                    if (string.Compare(name, arrHyperlink[count].ToString()) == 0) {
                        flag = 1;
                    }
                }

            }
            return flag.ToString();
        }

        /// <summary>
        /// Function to increments flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string SetHyperLinkFlag() {
            setHyperLinkFlag++;
            return setHyperLinkFlag.ToString();
        }

        /// <summary>
        /// Function which returns flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string GetHyperlinkFlag() {
            return setHyperLinkFlag.ToString();
        }

        /// <summary>
        /// Function to set flag for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string SetGetHyperLinkFlag() {
            setHyperLinkFlag = 0;
            return setHyperLinkFlag.ToString();
        }

        /// <summary>
        /// Function to increments flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string HyperlinkFlag() {
            hyperlinkFlag++;
            return hyperlinkFlag.ToString();
        }

        /// <summary>
        /// Function to check whether hyperlink is coming is different runs or not
        /// </summary>
        /// <returns></returns>
        public string TestRun() {
            testRun++;
            return testRun.ToString();
        }

        /// <summary>
        /// Function to get the flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string GetTestRun() {
            return testRun.ToString();
        }


        /// <summary>
        /// Function to set the flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public string SetTestRun() {
            testRun = 0;
            return testRun.ToString();
        }


        /// <summary>
        /// Function to store ID
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string StroreId(string id) {
            storeHyperId = id;
            return id;
        }

        /// <summary>
        /// Function to check whether given Id is there or not 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public int CheckId(string id) {
            if (string.Compare(id, storeHyperId) == 0)
                return 1;
            else
                return 0;
        }

        /// <summary>
        /// Function to increment flag for Hyperlink
        /// </summary>
        /// <returns></returns>
        public int SetHyperLink() {
            set++;
            return set;
        }

        /// <summary>
        /// Function to set flag for Hyperlink
        /// </summary>
        /// <returns></returns>
        public int GetHyperLink() {
            set = 0;
            return set;
        }

        /// <summary>
        /// Function to get flag value for Hyperlink
        /// </summary>
        /// <returns></returns>
        public int GetFlag() {
            return set;
        }

        /// <summary>
        /// Function to increment flag value for Bookmark
        /// </summary>
        /// <returns></returns>
        public int SetBookmark() {
            setbookmark++;
            return setbookmark;
        }

        /// <summary>
        /// Function to get flag value for Bookmark
        /// </summary>
        /// <returns></returns>
        public int GetBookmark() {
            return setbookmark;
        }

        /// <summary>
        /// Function to set flag value for Bookmark
        /// </summary>
        /// <returns></returns>
        public int AssingBookmark() {
            setbookmark = 0;
            return setbookmark;
        }

        /// <summary>
        /// Function to return special character
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string EscapeSpecial(string id) {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < id.Length; i++) {
                if ((id[i] >= '0' && id[i] <= '9') || (id[i] >= 'A' && id[i] <= 'z')) {
                    sb.Append(id[i]);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Function which returns whether the BookmarkEnd related to Abbreviation or Acronym in document.xml
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public String Book(String id) {
            String indicator = "false";
            if (id != "") {
                if (id.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                    indicator = "AbbrTrue";
                }
                if (id.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                    indicator = "AcrTrue";
                }
            }

            return indicator;

        }


        /// <summary>
        /// Function which returns whether the BookmarkEnd related to Abbreviation or Acronym in footnote.xml/endnote.xml
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public String BookFootnote(String id) {
            String indicator = "false";
            XmlNodeList node;

            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(footRelationshipType)) {
                footrelationship = searchRelation;
                break;
            }

            Uri footpartUri = PackUriHelper.ResolvePartUri(footrelationship.SourceUri, footrelationship.TargetUri);
            PackagePart footPartxml = pack.GetPart(footpartUri);

            XmlDocument doc = new XmlDocument();
            doc.Load(footPartxml.GetStream());

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", wordNamespace);

            node = doc.SelectNodes("//w:bookmarkStart[@w:id='" + id + "']", nsManager);

            if (node.Count == 0) {

                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri1 = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml1 = pack.GetPart(partUri1);

                foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(endRelationshipType)) {
                    endrelationship = searchRelation;
                    break;
                }

                Uri endpartUri = PackUriHelper.ResolvePartUri(endrelationship.SourceUri, endrelationship.TargetUri);
                PackagePart endPartxml = pack.GetPart(endpartUri);

                XmlDocument doc1 = new XmlDocument();
                doc1.Load(endPartxml.GetStream());

                NameTable nt1 = new NameTable();
                XmlNamespaceManager nsManager1 = new XmlNamespaceManager(nt1);
                nsManager1.AddNamespace("w", wordNamespace);

                node = doc1.SelectNodes("//w:bookmarkStart[@w:id='" + id + "']", nsManager1);
            }
            if (node.Count != 0) {
                if (node.Item(0).Attributes.Item(1).Value.StartsWith("Abbreviations", StringComparison.CurrentCulture)) {
                    indicator = "AbbrTrue";
                }
                if (node.Item(0).Attributes.Item(1).Value.StartsWith("Acronyms", StringComparison.CurrentCulture)) {
                    indicator = "AcrTrue";
                }
            } else
                indicator = "false";

            return indicator;

        }


        /// <summary>
        /// Function which returns Full form of an Abbreviation
        /// </summary>
        /// <param name="abbrName"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public String FullAbbr(String abbrName, String version) {
            String indicator = "";
            relationship = null;
            if (IsOffice2007Or2010(version)) {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(CustomRelationshipType)) {

                    customRelationship = searchRelation;

                    Uri CustomUri = PackUriHelper.ResolvePartUri(customRelationship.SourceUri, customRelationship.TargetUri);
                    PackagePart customPartxml = pack.GetPart(CustomUri);

                    XmlDocument doc = new XmlDocument();
                    doc.Load(customPartxml.GetStream());

                    if (doc.DocumentElement.NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        NameTable nt = new NameTable();

                        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                        nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                        XmlNodeList node = doc.SelectNodes("//a:Item[@AbbreviationName='" + abbrName + "']", nsManager);

                        if (node != null) {
                            if (node.Count != 0) {
                                if (node.Item(0).Attributes.Item(1).Value != "")
                                    indicator = node.Item(0).Attributes.Item(1).Value;
                            } else
                                indicator = "";
                            break;
                        }
                    }
                }
            }

            if (version == version2003 || version == versionXP) {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(customPropRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                if (relationship != null) {
                    Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                    PackagePart mainPartxml = pack.GetPart(partUri);

                    XmlDocument doc = new XmlDocument();
                    doc.Load(mainPartxml.GetStream());

                    XmlNodeList node = doc.FirstChild.NextSibling.ChildNodes;

                    if (node != null) {
                        for (int i = 0; i < node.Count; i++) {
                            if (node.Item(i).Attributes.Item(2).Value == abbrName) {
                                if (node.Item(i).FirstChild.InnerText != "") {
                                    String temp = node.Item(i).FirstChild.InnerText;
                                    String input = temp.Replace("$#$", "-");
                                    String[] strKey = input.Split('-');
                                    indicator = strKey[1];
                                } else
                                    indicator = "";
                            }
                        }
                    }
                }
            }
            return indicator;
        }


        /// <summary>
        /// Function which returns Full form of an Acronym
        /// </summary>
        /// <param name="acrName"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public String FullAcr(String acrName, String version) {
            String indicator = "";
            relationship = null;
            if (IsOffice2007Or2010(version)) {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                PackagePart mainPartxml = pack.GetPart(partUri);

                foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(CustomRelationshipType)) {
                    customRelationship = searchRelation;


                    Uri CustomUri = PackUriHelper.ResolvePartUri(customRelationship.SourceUri, customRelationship.TargetUri);
                    PackagePart customPartxml = pack.GetPart(CustomUri);

                    XmlDocument doc = new XmlDocument();
                    doc.Load(customPartxml.GetStream());

                    if (doc.DocumentElement.NamespaceURI == "http://Daisy-OpenXML/customxml") {
                        NameTable nt = new NameTable();

                        XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                        nsManager.AddNamespace("a", "http://Daisy-OpenXML/customxml");

                        XmlNodeList node = doc.SelectNodes("//a:Item[@AcronymName='" + acrName + "']", nsManager);

                        if (node.Count != 0) {
                            if (node.Item(0).Attributes.Item(1).Value != "")
                                indicator = node.Item(0).Attributes.Item(1).Value;
                        } else
                            indicator = "";
                        break;
                    }

                }
            }

            if (version == version2003 || version == versionXP) {
                foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(customPropRelationshipType)) {
                    relationship = searchRelation;
                    break;
                }

                if (relationship != null) {
                    Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
                    PackagePart mainPartxml = pack.GetPart(partUri);

                    XmlDocument doc = new XmlDocument();
                    doc.Load(mainPartxml.GetStream());

                    XmlNodeList node = doc.FirstChild.NextSibling.ChildNodes;
                    if (node != null) {
                        for (int i = 0; i < node.Count; i++) {
                            if (node.Item(i).Attributes.Item(2).Value == acrName) {
                                if (node.Item(i).FirstChild.InnerText != "") {
                                    String temp = node.Item(i).FirstChild.InnerText;
                                    String input = temp.Replace("$#$", "-");
                                    String[] strKey = input.Split('-');
                                    indicator = strKey[1];
                                } else
                                    indicator = "";
                            }
                        }
                    }
                }
            }
            return indicator;

        }


        /// <summary>
        /// Function used to set the Abbreviations/Acronyms flag
        /// </summary>
        /// <returns></returns>
        public String SetAbbrAcrFlag() {
            AbbrAcrflag = 1;
            return "1";
        }

        /// <summary>
        /// Function used to reset the Abbreviations/Acronyms flag
        /// </summary>
        /// <returns></returns>
        public String ReSetAbbrAcrFlag() {
            AbbrAcrflag = 0;
            return "0";
        }


        /// <summary>
        /// Function used to return the Abbreviations/Acronyms flag value
        /// </summary>
        /// <returns></returns>
        public int AbbrAcrFlag() {
            return AbbrAcrflag;
        }

        /// <summary>
        /// Function used to push an abbreviation to stack
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public String PushAbrAcr(String element) {
            abbrstackList.Push(element);
            return "1";
        }


        /// <summary>
        /// Function used to push the given element into stack
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public String PushAbrAcrpara(String element) {
            abbrparastackList.Push(element);
            return "1";
        }

        /// <summary>
        /// Function used to push the given element into stack
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public String PushAbrAcrhead(String element) {
            abbrheadstackList.Push(element);
            return "1";
        }

        /// <summary>
        /// Function used to peek an element from the stack
        /// </summary>
        /// <returns></returns>
        public String PeekAbrAcr() {
            String temp = abbrstackList.Pop();
            return temp;
        }

        /// <summary>
        /// Function used to peek an element from the stack
        /// </summary>
        /// <returns></returns>
        public String PeekAbrAcrpara() {
            String temp = abbrparastackList.Pop();
            return temp;
        }


        /// <summary>
        /// Function used to peek an element from the stack 
        /// </summary>
        /// <returns></returns>
        public String PeekAbrAcrhead() {
            String temp = abbrheadstackList.Pop();
            return temp;
        }

        /// <summary>
        /// Function which returns the count value
        /// </summary>
        /// <returns></returns>
        public String CountAbrAcr() {
            return abbrstackList.Count.ToString();
        }

        /// <summary>
        /// Function which returns the count value
        /// </summary>
        /// <returns></returns>
        public String CountAbrAcrpara() {
            return abbrparastackList.Count.ToString();
        }

        /// <summary>
        /// Function which returns the count value
        /// </summary>
        /// <returns></returns>
        public String CountAbrAcrhead() {
            return abbrheadstackList.Count.ToString();
        }


        /// <summary>
        /// Function to return Rowspan value of a table
        /// </summary>
        /// <returns></returns>
        public string Rowspan() {
            rowspan++;
            return rowspan.ToString();
        }

        /// <summary>
        /// Function to set Rowspan value of a table
        /// </summary>
        /// <returns></returns>
        public string SetRowspan() {
            rowspan = 0;
            return rowspan.ToString();
        }

        /// <summary>
        /// Function to return Rowspan value of a table
        /// </summary>
        /// <returns></returns>
        public string GetRowspan() {
            return rowspan.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string GetFlagRowspan() {
            flagRowspan++;
            return flagRowspan.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string SetFlagRowspan() {
            flagRowspan = 0;
            return flagRowspan.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ReturnFlagRowspan() {
            return flagRowspan.ToString();
        }

        /// <summary>
        /// Function to increment flag value for BDO
        /// </summary>
        /// <returns></returns>
        public String SetbdoFlag() {
            bdoflag++;
            return bdoflag.ToString();
        }

        /// <summary>
        /// Function to set flag value for BDO
        /// </summary>
        /// <returns></returns>
        public String reSetbdoFlag() {
            bdoflag = 0;
            return bdoflag.ToString();
        }


        /// <summary>
        /// Function to increment flag value for RTL 
        /// </summary>
        /// <returns></returns>
        public String SetRtlFlag() {
            rtlFlag++;
            return rtlFlag.ToString();
        }

        /// <summary>
        /// Function to set flag value for RTL
        /// </summary>
        /// <returns></returns>
        public String ResetRtlFlag() {
            rtlFlag = 0;
            return rtlFlag.ToString();
        }

        /// <summary>
        /// Function to increment flag value for caption 
        /// </summary>
        /// <returns></returns>
        public String SetcaptionFlag() {
            captionFlag++;
            return captionFlag.ToString();
        }

        /// <summary>
        /// Function to set flag value for caption
        /// </summary>
        /// <returns></returns>
        public String reSetcaptionFlag() {
            captionFlag = 0;
            return captionFlag.ToString();
        }

        /// <summary>
        /// Function to increment flag value for Cover page 
        /// </summary>
        /// <returns></returns>
        public string CheckCaverPage() {
            checkCverpage++;
            return checkCverpage.ToString();
        }

        /// <summary>
        /// Function to increment flag value for cover page
        /// </summary>
        /// <returns></returns>
        public string CodeFlag() {
            codeFlag++;
            return codeFlag.ToString();
        }


        /// <summary>
        /// Function to returns flag value for cover page
        /// </summary>
        /// <returns></returns>
        public string GetCodeFlag() {
            return codeFlag.ToString();
        }

        /*Function to set flag value for cover page */
        public string InitializeCodeFlag() {
            codeFlag = 0;
            return codeFlag.ToString();
        }

        /*Function to set flag value for Heading */
        public int ListHeadingFlag() {
            return listHeadingFlag;
        }

        /*Function used to push the given element into stack*/
        public String PushListHeading(String element) {
            listHeadingstackList.Push(element);
            return "1";
        }

        /*Function used to peek an element from the stack*/
        public String PeekListHeading() {
            String temp = "";
            while (listHeadingstackList.Count > 0) {
                temp = listHeadingstackList.Pop() + temp;
            }
            return temp;
        }

        /*Function used to set the Abbreviations/Acronyms flag*/
        public String SetListHeadingFlag() {
            listHeadingFlag = 1;
            return "1";
        }

        /*Function used to reset the Abbreviations/Acronyms flag*/
        public String ReSetListHeadingFlag() {
            listHeadingFlag = 0;
            return "0";
        }

        #region Different Langauage

        /*Function used to compare two Headings info*/
        public String CompareHeading(String strA, String strB) {
            String value = "0";
            if (strA != "" && strB != "") {
                if (strA.Substring(0, strA.Length - 1).Equals(strB.Substring(0, strB.Length - 1)))
                    value = "1";
            }
            return value;

        }

        #endregion

        # region LineNum

        /*Function used to set the Linenumber flag*/
        public String Setlinenumflag() {
            linenumflag = 1;
            return "1";
        }

        /*Function used to get the Linenumber flag value*/
        public int Getlinenumflag() {
            return linenumflag;
        }

        /*Function used to reset the Linenumber flag*/
        public String Resetlinenumflag() {
            linenumflag = 0;
            return "0";
        }

        #endregion

        #region MultipleOOXML

        /*Function used to get the masterSub flag*/
        public int ListMasterSubFlag() {
            return listMasterSubFlag;
        }

        /*Function used to push the given element into stack*/
        public String PushMasterSubdoc(String element) {
            masteSubstackList.Push(element);
            return "1";
        }

        /*Function used to peek an element from the stack*/
        public String PeekMasterSubdoc() {
            String temp = "";
            while (masteSubstackList.Count > 0) {
                temp = masteSubstackList.Pop() + temp;
            }
            return temp;
        }

        /*Function used to set the MultipleOOXML flag*/
        public String MasterSubSetFlag() {
            listMasterSubFlag = 1;
            return "1";
        }

        /*Function used to reset the MultipleOOXML flag*/
        public String MasterSubResetFlag() {
            listMasterSubFlag = 0;
            return "0";
        }

        /*Function used to close levels for MultipleOOXML */
        public String ClosingMasterSub(String value) {
            String output = "";
            for (int i = Convert.ToInt16(value); i >= 1; i--) {
                output = output + "</level" + i + ">";
            }

            output = "<p/>" + output;
            return output;
        }

        /*Function used to open levels for MultipleOOXML */
        public String OpenMasterSub(String value) {
            String output = "";
            for (int i = 1; i <= Convert.ToInt16(value); i++) {
                output = output + "<level" + i + ">";
            }

            output = output + "<p/>";

            return output;
        }
        #endregion

        #region List

        public String DiffLevel(String opLvl, String PeekLvl) {
            return Convert.ToString(Convert.ToInt16(opLvl) - Convert.ToInt16(PeekLvl));
        }

        public String ReduceOne(String opLvl) {
            return Convert.ToString(Convert.ToInt16(opLvl) - 1);
        }

        #endregion

        #region Lists Numbering

        /*Function which increments the counter value*/
        public void Increment(int rec, int pkLvl, String numId) {
            for (int i = 0; i <= rec; i++) {
                ListPush(i);
                if (i > pkLvl)
                    IncrementListCounters(i.ToString(), numId);

            }
        }

        public String TextList(String numFormat, String lvlText, String numId, String iLvl) {
            String text = "";
            int index = lvlText.IndexOf('%');
            if (index < 0) {
                if (numFormat == "bullet") {
                    text = bulletChar[Convert.ToInt16(iLvl)].ToString();
                } else if (numFormat == "none") {
                    text = "";
                } else if (numFormat == "") {
                    text = "";
                } else {
                    text = lvlText;
                }
                text = text + " ";
            } else {
                text = lvlText.Substring(lvlText.IndexOf('%'));
                AbstractFormat(numId, iLvl);
                int cntSymbol = 0;

                for (int j = 0; j < text.Length; j++) {
                    if (text[j].ToString() == "%")
                        cntSymbol++;
                }

                string[] chr = new string[text.Length];
                Int16 cntCopy;
                if (cntSymbol == 1)
                    cntCopy = NonZeroListCounter(numId);
                else
                    cntCopy = 0;
                for (int i = 0; i < text.Length; i++) {
                    chr.SetValue(text[i].ToString(), i);
                    if (i > 0) {
                        if (chr.GetValue(i - 1).ToString() == "%") {
                            if (numFormat == "decimalZero") {
                                if (cntSymbol == 1) {
                                    String valList = SpecificFormat(cntCopy.ToString(), numFormat, iLvl);
                                    chr.SetValue("0" + valList, i);
                                } else {
                                    if (cntCopy > 0) {
                                        String valList = SpecificFormat(((ArrayList)listCounters["List" + numId])[cntCopy].ToString(), numFormat, iLvl);
                                        if (Convert.ToInt16(valList) < 9)
                                            chr.SetValue("0" + valList, i);
                                        else
                                            chr.SetValue(valList, i);
                                    } else {
                                        String valList = SpecificFormat(((ArrayList)listCounters["List" + numId])[cntCopy].ToString(), numFormat, iLvl);
                                        chr.SetValue(valList, i);
                                    }
                                }
                            } else {
                                if (cntSymbol == 1) {
                                    String valList = SpecificFormat(cntCopy.ToString(), numFormat, iLvl);
                                    chr.SetValue(valList, i);
                                } else {
                                    String valList = SpecificFormat(((ArrayList)listCounters["List" + numId])[cntCopy].ToString(), absFormat[cntCopy].ToString(), iLvl);
                                    chr.SetValue(valList, i);
                                }
                            }
                            cntCopy++;
                        }
                    }
                }
                String temp = "";
                for (int i = 0; i < chr.Length; i++) {
                    temp = temp + chr[i];
                }
                temp = temp.Replace("%", "");
                text = lvlText.Substring(0, lvlText.IndexOf('%')) + temp;
            }

            return text + " ";
        }

        public String SpecificFormat(String lvlText, String numFormat, String iLvl) {
            String tempString = "";
            if (numFormat == "decimal" || numFormat == "decimalZero") {
                tempString = lvlText;
            } else if (numFormat == "lowerLetter") {
                tempString = PageNumLowerAlphabet(lvlText);
            } else if (numFormat == "upperLetter") {
                tempString = PageNumUpperAlphabet(lvlText);
            } else if (numFormat == "upperRoman") {
                tempString = PageNumUpperRoman(lvlText);
            } else if (numFormat == "lowerRoman") {
                tempString = PageNumLowerRoman(lvlText);
            } else if (numFormat == "bullet") {
                tempString = bulletChar[Convert.ToInt16(iLvl)].ToString();
            } else if (numFormat == "none") {
                tempString = "";
            } else {
                tempString = lvlText;
            }
            return tempString;
        }

        public void AbstractFormat(String numId, String iLvl) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(numberRelationshipType)) {
                numberRelationship = searchRelation;
                break;
            }

            Uri numberPartUri = PackUriHelper.ResolvePartUri(numberRelationship.SourceUri, numberRelationship.TargetUri);
            PackagePart numberPartxml = pack.GetPart(numberPartUri);

            XmlDocument doc = new XmlDocument();
            doc.Load(numberPartxml.GetStream());

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", docNamespace);


            XmlNodeList list = doc.SelectNodes("w:numbering/w:num[@w:numId=" + numId + "]/w:abstractNumId", nsManager);
            if (list.Count != 0) {
                String absNumid = list[0].Attributes[0].Value;
                XmlNodeList listAbs = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + absNumid + "]/w:lvl", nsManager);
                if (listAbs.Count != 0) {
                    for (int j = 0; j < listAbs.Count; j++)
                        absFormat[j] = listAbs[j].SelectSingleNode("w:numFmt", nsManager).Attributes[0].Value;
                }
            }
        }

        public String IsList(String numId) {
            if (listCounters.ContainsKey("List" + numId))
                return "ListTrue";
            else if (headingCounters.ContainsKey("List" + numId))
                return "HeadTrue";
            else
                return numId;

        }

        public void IncrementListCounters(String iLvl, String numId) {
            if (startItem[Convert.ToInt16(iLvl)].ToString() == "")
                startItem[Convert.ToInt16(iLvl)] = "0";

            if (((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)].ToString() == DaisyClass.emptyListCounter) {
                if (Convert.ToInt16(startItem[Convert.ToInt16(iLvl)]) != 1)
                    ((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)] = startItem[Convert.ToInt16(iLvl)].ToString();
                else
                    ((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)] = 1;
            } else
                ((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)] = Convert.ToInt16(((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)]) + 1;

            for (int i = Convert.ToInt16(iLvl) + 1; i < ((ArrayList)listCounters["List" + numId]).Count; i++) {
                ((ArrayList)listCounters["List" + numId])[i] = DaisyClass.emptyListCounter;
            }

            for (int i = 0; i <= Convert.ToInt16(iLvl); i++) {
                if (((ArrayList)listCounters["List" + numId])[Convert.ToInt16(i)].ToString() == DaisyClass.emptyListCounter)
                    ((ArrayList)listCounters["List" + numId])[Convert.ToInt16(i)] = startItem[i].ToString();
            }

        }

        public int GetListCounter(string iLvl, string numId) {
            string listCounter = ((ArrayList)listCounters["List" + numId])[Convert.ToInt16(iLvl)].ToString();
            if (listCounter == DaisyClass.emptyListCounter)
                return 1;
            else
                return int.Parse(listCounter) + 1;
        }

        public Int16 NonZeroListCounter(String numId) {
            Int16 nonZeroValue = -1;
            Int16 i = 0;
            for (i = 0; i < ((ArrayList)listCounters["List" + numId]).Count; i++) {
                if (((ArrayList)listCounters["List" + numId])[i].ToString() == DaisyClass.emptyListCounter) {
                    nonZeroValue = Convert.ToInt16(((ArrayList)listCounters["List" + numId])[i - 1].ToString());
                    break;
                }
            }
            if (nonZeroValue == -1)
                nonZeroValue = Convert.ToInt16(((ArrayList)listCounters["List" + numId])[8].ToString());

            return nonZeroValue;
        }

        public String ResetListCounters() {
            foreach (DictionaryEntry myEntry in listCounters) {
                for (int i = 0; i < 9; i++)
                    ((ArrayList)myEntry.Value)[i] = DaisyClass.emptyListCounter;
            }
            listCounters.Clear();
            return "1";
        }

        public String GetCurrentNumID() {
            return prevNumId;
        }

        public String CheckNumID(String numId) {
            if (prevNumId != numId) {
                prevNumId = numId;
                return "True";
            } else {
                return "False";
            }
        }

        public string StartNewListCounter(String numId) {
            if (!listCounters.ContainsKey("List" + numId)) {
                listCounters.Add("List" + numId, new ArrayList(9));

                for (int i = 0; i < 9; i++)
                    ((ArrayList)listCounters["List" + numId]).Add(DaisyClass.emptyListCounter);
            }

            return "1";
        }

        /*Function which decrements the counter value*/
        public int DecrementStart(int rec) {
            rec = rec - 1;
            return rec;
        }

        /*Function which decrements the counter value*/
        public String StartString(int iLvl, int strtItem) {
            startItem[iLvl] = strtItem.ToString();
            return "1";
        }

        #endregion

        #region Heading Numbering

        public void StoreHeadingPart(String headingInfo) {
            this.headingInfo = headingInfo;
        }

        public String RetrieveHeadingPart() {
            return this.headingInfo;
        }

        public String TextHeading(String numFormat, String lvlText, String numId, String iLvl) {

            String text = "";
            int index = lvlText.IndexOf('%');
            if (index < 0) {
                if (numFormat == "bullet") {
                    text = bulletChar[Convert.ToInt16(iLvl)].ToString();
                } else if (numFormat == "none") {
                    text = "";
                } else if (numFormat == "") {
                    text = "";
                } else {
                    text = lvlText;
                }
                text = text + " ";
            } else {
                text = lvlText.Substring(lvlText.IndexOf('%'));
                AbstractFormat(numId, iLvl);
                int cntSymbol = 0;

                for (int j = 0; j < text.Length; j++) {
                    if (text[j].ToString() == "%")
                        cntSymbol++;
                }

                string[] chr = new string[text.Length];
                Int16 cntCopy;
                if (cntSymbol == 1)
                    cntCopy = NonZeroHeadingCounter(numId);
                else
                    cntCopy = 0;
                for (int i = 0; i < text.Length; i++) {
                    chr.SetValue(text[i].ToString(), i);
                    if (i > 0) {
                        if (chr.GetValue(i - 1).ToString() == "%") {
                            if (numFormat == "decimalZero") {
                                if (cntSymbol == 1) {
                                    String valList = SpecificFormat(cntCopy.ToString(), numFormat, iLvl);
                                    chr.SetValue("0" + valList, i);
                                } else {
                                    if (cntCopy > 0) {
                                        String valList = SpecificFormat(((ArrayList)headingCounters["List" + numId])[cntCopy].ToString(), numFormat, iLvl);
                                        if (Convert.ToInt16(valList) < 9)
                                            chr.SetValue("0" + valList, i);
                                        else
                                            chr.SetValue(valList, i);
                                    } else {
                                        String valList = SpecificFormat(((ArrayList)headingCounters["List" + numId])[cntCopy].ToString(), numFormat, iLvl);
                                        chr.SetValue(valList, i);
                                    }
                                }
                            } else {
                                if (cntSymbol == 1) {
                                    String valList = SpecificFormat(cntCopy.ToString(), numFormat, iLvl);
                                    chr.SetValue(valList, i);
                                } else {
                                    String valList = SpecificFormat(((ArrayList)headingCounters["List" + numId])[cntCopy].ToString(), absFormat[cntCopy].ToString(), iLvl);
                                    chr.SetValue(valList, i);
                                }
                            }
                            cntCopy++;
                        }
                    }
                }
                String temp = "";
                for (int i = 0; i < chr.Length; i++) {
                    temp = temp + chr[i];
                }
                temp = temp.Replace("%", "");
                text = lvlText.Substring(0, lvlText.IndexOf('%')) + temp;
            }

            return text + " ";
        }

        public void BulletChar() {
            bulletChar.Add("\u2605");
            bulletChar.Add("\u25B6");
            bulletChar.Add("\u25A3");
            bulletChar.Add("\u25CF");
            bulletChar.Add("\u25C6");
            bulletChar.Add("\u25CB");
            bulletChar.Add("\u25B2");
            bulletChar.Add("\u25C8");
            bulletChar.Add("\u25C7");
        }

        public void IncrementHeadingCounters(String iLvl, String numId, String absId) {
            if (numId != "" && absId != "") {
                String tempId = "";
                tempId = CheckAbstCounter(numId, absId);

                switch (((ArrayList)startHeadingItem["List" + tempId])[Convert.ToInt16(iLvl)].ToString()) {
                    case "Inc":

                        switch (((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)].ToString()) {
                            case DaisyClass.emptyListCounter:
                                ((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)] = 1;
                                break;
                            default:
                                ((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)] = Convert.ToInt16(((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)]) + 1;
                                break;
                        }
                        break;

                    default:
                        ((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)] = ((ArrayList)startHeadingItem["List" + tempId])[Convert.ToInt16(iLvl)].ToString();
                        break;
                }

                for (int i = Convert.ToInt16(iLvl) + 1; i < ((ArrayList)headingCounters["List" + numId]).Count; i++) {
                    ((ArrayList)headingCounters["List" + numId])[i] = DaisyClass.emptyListCounter;
                }
            }
        }

        public String CheckAbstCounter(String numId, String absId) {
            String tempId = "";
            if (!startHeadingItem.ContainsKey("List" + numId)) {
                if (absId == "")
                    absId = baseAbsId;

                if (baseAbsId == absId)
                    tempId = baseNumId;
            } else
                tempId = numId;

            return tempId;
        }

        public Int16 NonZeroHeadingCounter(String numId) {
            Int16 nonZeroValue = -1;
            Int16 i = 0;
            for (i = 0; i < ((ArrayList)headingCounters["List" + numId]).Count; i++) {
                if (((ArrayList)headingCounters["List" + numId])[i].ToString() == DaisyClass.emptyListCounter) {
                    nonZeroValue = Convert.ToInt16(((ArrayList)headingCounters["List" + numId])[i - 1].ToString());
                    break;
                }
            }
            if (nonZeroValue == -1)
                nonZeroValue = Convert.ToInt16(((ArrayList)headingCounters["List" + numId])[8].ToString());

            return nonZeroValue;

        }

        public String ResetHeadingCounters() {
            foreach (DictionaryEntry myEntry in headingCounters) {
                for (int i = 0; i < 9; i++)
                    ((ArrayList)myEntry.Value)[i] = "";
            }
            headingCounters.Clear();
            return "1";
        }

        public String CheckHeadingNumID(String numId) {
            String diffNumId = "";
            if (prevHeadNumId != numId) {
                if (prevHeadNumId != "") {
                    diffNumId = "True";
                    prevHeadNumId = numId;
                } else {
                    diffNumId = "True";
                    prevHeadNumId = numId;
                }
            } else {
                diffNumId = "False";
                prevHeadNumId = numId;
            }

            return diffNumId;
        }

        public String StartNewHeadingCounter(String numId, String absId) {
            if (headingCounters.Count == 0) {
                baseNumId = numId;
                baseAbsId = absId;
            }

            if (!headingCounters.ContainsKey("List" + numId)) {
                headingCounters.Add("List" + numId, new ArrayList(9));

                for (int i = 0; i < 9; i++)
                    ((ArrayList)headingCounters["List" + numId]).Add(DaisyClass.emptyListCounter);
            }

            return "1";
        }

        public String StartHeadingValueCtr(String numId, String absId) {
            if (!startHeadingItem.ContainsKey("List" + numId)) {
                if (baseAbsId != absId || !startHeadingItem.ContainsKey("List" + baseNumId)) {
                    startHeadingItem.Add("List" + numId, new ArrayList(9));

                    for (int i = 0; i < 9; i++)
                        ((ArrayList)startHeadingItem["List" + numId]).Add("");
                }
            }

            return "1";
        }

        public String StartHeadingNewCtr(String numId, String absId) {
            if (!startHeadingItem.ContainsKey("List" + numId)) {
                if (baseAbsId != absId || !startHeadingItem.ContainsKey("List" + baseNumId)) {
                    startHeadingItem.Add("List" + numId, new ArrayList(9));

                    for (int i = 0; i < 9; i++)
                        ((ArrayList)startHeadingItem["List" + numId]).Add("");
                }
            }

            return "1";
        }

        public string CopyToBaseCounter(String numId) {
            if (numId != baseNumId && numId != "" && numId != "0") {
                for (int i = 0; i < 9; i++) {
                    ((ArrayList)headingCounters["List" + baseNumId])[i] = ((ArrayList)headingCounters["List" + numId])[i];
                }

                return "1";
            } else
                return "2";

        }

        public string CopyToCurrCounter(String numId) {
            if (numId != baseNumId && numId != "" && numId != "0") {
                for (int i = 0; i < 9; i++) {
                    if (((ArrayList)headingCounters["List" + baseNumId])[i].ToString() != DaisyClass.emptyListCounter)
                        ((ArrayList)headingCounters["List" + numId])[i] = ((ArrayList)headingCounters["List" + baseNumId])[i];
                }
                return "1";
            } else
                return "2";

        }

        /*Function which decrements the counter value*/
        public String StartHeadingString(int iLvl, String strtItem, String numId, String absId, String location, String overRide) {
            if (strtItem == "")
                strtItem = "0";

            Int16 overrideFlag = CheckOverRideFlag(numId);

            if (baseAbsId == absId)
                numId = baseNumId;

            if (iLvl == 0) {
                for (int i = 1; i < 9; i++)
                    ((ArrayList)startHeadingItem["List" + numId])[i] = "";
            }

            int val = iLvl + 1;
            for (int i = val; i < 9; i++)
                ((ArrayList)startHeadingItem["List" + numId])[i] = "";

            if (overRide == "Yes" && location == "Document" && overrideFlag == 1) {
                ((ArrayList)startHeadingItem["List" + numId])[iLvl] = strtItem;
            } else {
                if (((ArrayList)startHeadingItem["List" + numId])[iLvl].ToString() == "") {
                    if ((baseAbsId == absId && baseNumId != numId && overRide == "No"))
                        ((ArrayList)startHeadingItem["List" + numId])[iLvl] = "Inc";

                    else
                        ((ArrayList)startHeadingItem["List" + numId])[iLvl] = strtItem;
                } else {
                    ((ArrayList)startHeadingItem["List" + numId])[iLvl] = "Inc";
                }
            }
            return "0";
        }

        public Int16 CheckOverRideFlag(String numId) {
            Int16 overrideFlag = 0;
            if (!OverideNumList.Contains(numId)) {
                OverideNumList.Add(numId);
                overrideFlag = 1;
            } else {
                overrideFlag = 0;
            }
            return overrideFlag;
        }

        public String AddCurrHeadId(String currId) {
            if (currId != "0" && currId != "") {
                if (prevHeadId[0].ToString() == "") {
                    prevHeadId[0] = currId;
                    prevHeadId[1] = currId;
                } else {
                    prevHeadId[0] = prevHeadId[1];
                    prevHeadId[1] = currId;
                }
            }
            return "1";
        }

        public String AddCurrHeadLevel(String currLvl, String location, String absId) {
            String diffValue = "" + "|" + prevHeadLvl;

            if (absId != "") {
                if (prevHeadLvl == "") {
                    if (Convert.ToInt16(currLvl) > 0) {
                        for (int i = 0; i < Convert.ToInt16(currLvl); i++) {
                            StartHeadingValueCtr(baseNumId, baseAbsId);
                            StartHeadingString(i, "1", baseNumId, baseAbsId, "Document", "No");
                            IncrementHeadingCounters(i.ToString(), baseNumId, baseAbsId);
                        }
                    }
                    prevHeadLvl = currLvl;
                    diffValue = "" + "|" + prevHeadLvl;
                } else {
                    if (Convert.ToInt16(currLvl) - Convert.ToInt16(prevHeadLvl) > 1) {
                        diffValue = Convert.ToInt16(currLvl) - Convert.ToInt16(prevHeadLvl) + "|" + prevHeadLvl;
                        NumberHeadings(Convert.ToInt16(prevHeadLvl), Convert.ToInt16(currLvl), location, absId);
                    }

                    prevHeadLvl = currLvl;
                }
            }
            return diffValue;
        }

        public void NumberHeadings(int prevHeadLvl, int currLvl, String location, String absId) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(numberRelationshipType)) {
                numberRelationship = searchRelation;
                break;
            }

            Uri numberPartUri = PackUriHelper.ResolvePartUri(numberRelationship.SourceUri, numberRelationship.TargetUri);
            PackagePart numberPartxml = pack.GetPart(numberPartUri);

            XmlDocument doc = new XmlDocument();
            doc.Load(numberPartxml.GetStream());

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", docNamespace);

            StartNewHeadingCounter(prevHeadId[1].ToString(), absId);
            CopyToCurrCounter(prevHeadId[1].ToString());

            int val = prevHeadLvl + 1;
            for (int i = val; i <= currLvl; i++) {
                XmlNodeList listDel = doc.SelectNodes("w:numbering/w:num[@w:numId=" + prevHeadId[0] + "]/w:lvlOverride[@w:ilvl=" + i + "]/w:startOverride", nsManager);
                if (listDel.Count != 0) {
                    StartHeadingValueCtr(prevHeadId[1].ToString(), absId);
                    String tempId = "";
                    tempId = CheckAbstCounter(prevHeadId[1].ToString(), absId);

                    String valAbs = listDel[0].Attributes[0].Value;

                    if (valAbs == "")
                        valAbs = "0";

                    if (i == currLvl) {
                        if (location == "Style") {
                            ((ArrayList)startHeadingItem["List" + tempId])[i] = Convert.ToInt16(valAbs) - 1;
                            ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = Convert.ToInt16(valAbs) - 1;
                        } else {
                            ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                            ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                        }
                    } else {
                        ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                        ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                    }


                } else {
                    XmlNodeList listAbsDel = doc.SelectNodes("w:numbering/w:num[@w:numId=" + prevHeadId[0] + "]/w:abstractNumId", nsManager);
                    XmlNodeList list = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + listAbsDel[0].Attributes[0].Value + "]/w:lvl[@w:ilvl=" + i + "]/w:start", nsManager);

                    StartHeadingValueCtr(prevHeadId[1].ToString(), absId);
                    String tempId = "";
                    tempId = CheckAbstCounter(prevHeadId[1].ToString(), absId);

                    if (list.Count != 0) {
                        String valAbs = list[0].Attributes[0].Value;

                        if (valAbs == "")
                            valAbs = "0";

                        if (i == currLvl) {
                            if (location == "Style") {
                                ((ArrayList)startHeadingItem["List" + tempId])[i] = Convert.ToInt16(valAbs) - 1;
                                ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = Convert.ToInt16(valAbs) - 1;
                            } else {
                                ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                                ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                            }
                        } else {
                            ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                            ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                        }
                    } else {
                        String valAbs = "0";

                        if (i == currLvl) {
                            if (location == "Style") {
                                ((ArrayList)startHeadingItem["List" + tempId])[i] = Convert.ToInt16(valAbs) - 1;
                                ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = Convert.ToInt16(valAbs) - 1;
                            } else {
                                ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                                ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                            }
                        } else {
                            ((ArrayList)startHeadingItem["List" + tempId])[i] = valAbs;
                            ((ArrayList)headingCounters["List" + prevHeadId[1]])[Convert.ToInt16(i)] = valAbs;
                        }
                    }
                }

            }
        }

        #endregion

        #region Shapes and Objects

        public String GenerateObjectId() {
            objectId++;
            return objectId.ToString();
        }

        public String GetObjectId() {
            return objectId.ToString();
        }

        // Object in destination folder

        public String Object(string inNum) {

            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);
            imgRelationship = mainPartxml.GetRelationship(inNum);
            Uri imgpartUri = PackUriHelper.ResolvePartUri(imgRelationship.SourceUri, imgRelationship.TargetUri);
            PackagePart objPartxml = pack.GetPart(imgpartUri);
            Stream stream = null;
            stream = objPartxml.GetStream();
            Int16 length = Int16.Parse(objPartxml.Uri.ToString().LastIndexOf('/').ToString());
            String strImgName = objPartxml.Uri.ToString().Substring(length + 1);
            String str = outputFilename.Insert(outputFilename.Length, "\\");
            String final = str.Insert(str.Length, strImgName);
            if (!strImgName.EndsWith(".bin")) {
                FileStream objectFileStream = new FileStream(final, FileMode.Create, FileAccess.Write);
                int Length = 256;
                Byte[] buffer = new Byte[Length];
                int bytesRead = stream.Read(buffer, 0, Length);
                // write the required bytes
                while (bytesRead > 0) {
                    objectFileStream.Write(buffer, 0, bytesRead);
                    bytesRead = stream.Read(buffer, 0, Length);
                }
                stream.Close();
                objectFileStream.Close();
            }
            return strImgName;
        }

        #endregion

        #region Updated Lists

        public String Val() {
            return absVal;
        }

        public String NumFormat() {
            return numFormat;
        }

        public String LvlText() {
            return lvlText;
        }

        public String LStartOverride() {
            return lStartOverride;
        }

        public String LStart() {
            return lStart;
        }

        public void GeneralList(String iLvl, String numId) {
            foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
                relationship = searchRelation;
                break;
            }

            Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
            PackagePart mainPartxml = pack.GetPart(partUri);

            foreach (PackageRelationship searchRelation in mainPartxml.GetRelationshipsByType(numberRelationshipType)) {
                numberRelationship = searchRelation;
                break;
            }

            Uri numberPartUri = PackUriHelper.ResolvePartUri(numberRelationship.SourceUri, numberRelationship.TargetUri);
            PackagePart numberPartxml = pack.GetPart(numberPartUri);

            XmlDocument doc = new XmlDocument();
            doc.Load(numberPartxml.GetStream());

            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", docNamespace);

            XmlNodeList list = doc.SelectNodes("w:numbering/w:num[@w:numId=" + numId + "]/w:abstractNumId", nsManager);
            if (list != null) {
                if (list.Count != 0) {
                    if (list.Item(0).Attributes[0].Value != "")
                        absVal = list.Item(0).Attributes[0].Value;
                } else
                    absVal = "";
            }


            XmlNodeList numFormatNode = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + absVal + "]/w:lvl[@w:ilvl=" + iLvl + "]/w:numFmt", nsManager);
            if (numFormatNode != null) {
                if (numFormatNode.Count != 0) {
                    if (numFormatNode.Item(0).Attributes[0].Value != "")
                        numFormat = numFormatNode.Item(0).Attributes[0].Value;
                } else
                    numFormat = "";
            }

            XmlNodeList lvlTextNode = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + absVal + "]/w:lvl[@w:ilvl=" + iLvl + "]/w:lvlText", nsManager);
            if (lvlTextNode != null) {
                if (lvlTextNode.Count != 0) {
                    if (lvlTextNode.Item(0).Attributes[0].Value != "")
                        lvlText = lvlTextNode.Item(0).Attributes[0].Value;
                } else
                    lvlText = "";
            }

            XmlNodeList lStartOverrideNode = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + absVal + "]/w:lvl[@w:ilvl=" + iLvl + "]/w:start", nsManager);
            if (lStartOverrideNode != null) {
                if (lStartOverrideNode.Count != 0) {
                    if (lStartOverrideNode.Item(0).Attributes[0].Value != "")
                        lStartOverride = lStartOverrideNode.Item(0).Attributes[0].Value;
                } else
                    lStartOverride = "";
            }

            XmlNodeList lStartNode = doc.SelectNodes("w:numbering/w:abstractNum[@w:abstractNumId=" + absVal + "]/w:lvl[@w:ilvl=" + iLvl + "]/w:start", nsManager);
            if (lStartNode != null) {
                if (lStartNode.Count != 0) {
                    if (lStartNode.Item(0).Attributes[0].Value != "")
                        lStart = lStartNode.Item(0).Attributes[0].Value;
                } else
                    lStart = "";
            }
        }


        #endregion

        #region URI Images

        public String UriEscape(String imgName) {
            String imgNameRet;
            imgNameRet = HttpUtility.UrlEncode(imgName);
            if (imgNameRet.Contains("+"))
                imgNameRet = imgNameRet.Replace("+", "%20");
            return imgNameRet;
        }
        #endregion

        public String DocPropSubject() {
            String strCreator = pack.PackageProperties.Subject;
            if (strCreator != "" && strCreator != null)
                return strCreator;
            else
                return "";
        }

        public String DocPropDescription() {
            String strCreator = pack.PackageProperties.Description;
            if (strCreator != "" && strCreator != null)
                return strCreator;
            else
                return "";
        }

        /*Function used to set the Abbreviations/Acronyms flag*/
        public String SetListFlag() {
            listflag = 1;
            return "1";
        }

        /*Function used to reset the Abbreviations/Acronyms flag*/
        public String ReSetListFlag() {
            listflag = 0;
            return "0";
        }

        /*Function used to return the Abbreviations/Acronyms flag value*/
        public int ListFlag() {
            return listflag;
        }

        #region help methods

        private bool IsOffice2007Or2010(string version) {
            return version == version2007 || version == version2010;
        }

        #endregion

        #region matter type processing (Frontmatter (DAISY)/Bodymatter (DAISY)/Rearematter (DAISY) )

        PageStylesValidator _pageStylesValidator = new PageStylesValidator();
        List<PageStyle> _currentParagraphStylse = new List<PageStyle>();
        private StringBuilder _pageStylesErrors = new StringBuilder();
        private bool _isAnyPageStyleApplied = false;
        private bool _isRearmatterStyleApplied = false;

        public string PushPageStyle(string pageStyle) {
            _isAnyPageStyleApplied = true;
            var pushingStyle = PageStylesValidator.GetPageStyle(pageStyle);
            if (pushingStyle == PageStyle.Rearmatter)
                _isRearmatterStyleApplied = true;
            _currentParagraphStylse.Add(pushingStyle);
            return pageStyle;
        }

        public void IncrementCheckingParagraph() {
            var result = _pageStylesValidator.ValidateParagraph(_currentParagraphStylse);
            if (!result.IsValid)
                _pageStylesErrors.Append(result.ErrorMessage + Environment.NewLine);
            _currentParagraphStylse = new List<PageStyle>();
        }

        public string IsInvalidPageStylesSequence() {
            var result = _pageStylesValidator.ValidateLastStyle();
            if (!result.IsValid)
                _pageStylesErrors.Append(result.ErrorMessage);

            if (_pageStylesErrors.Length > 0)
                return true.ToString().ToLower();
            else
                return false.ToString().ToLower();
        }

        public string GetPageStylesErrors() {
            return _pageStylesErrors.ToString();
        }

        public string IsAnyPageStyleApplied() {
            return _isAnyPageStyleApplied.ToString().ToLower();
        }

        private string _currentMatterType = string.Empty;

        public void SetCurrentMatterType(string matterType) {
            _currentMatterType = matterType;
        }

        public string GetCurrentMatterType() {
            if (!_isAnyPageStyleApplied)
                return "Bodymatter";
            if (string.IsNullOrEmpty(_currentMatterType))
                return "Frontmatter";
            return _currentMatterType;
        }

        public void ResetCurrentMatterType() {
            _currentMatterType = string.Empty;
        }

        public string IsRearmatterStyleApplied() {
            return _isRearmatterStyleApplied.ToString().ToLower();
        }

        private bool _isFirstParagraph = true;

        public void ResetIsFirstParagraph() {
            _isFirstParagraph = true;
        }

        public void FirstParagraphProcessed() {
            _isFirstParagraph = false;
        }

        public string IsFirstParagraph() {
            return _isFirstParagraph.ToString().ToLower();
        }

        #endregion
    }
}