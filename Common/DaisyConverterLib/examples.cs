
using System;
using System.IO.Packaging;
using System.Collections;
using System.Collections.Generic; // For Stacks
using System.IO;
using System.Xml;
using System.Text;

namespace org.daisy.samples {

    /// <summary>
    /// Je reprend ici des exemples types de code utilisé dans les xslt mais définis 
    /// dans une classe c# "DaisyClass" du plugin word
    /// </summary>
    class DaisyClassSample {

        #region Variables et constantes

        // Je reprend toutes les constantes et variables pour l'exemple mais tu peux regarder directement les fonctions
        // et revenir sur les variables si besoin.
        // j'ai commencé a faire un peu de tri, parcequ'elles sont complètement en vrac dans le fichier d'origine.

        /* File "relationships" (types used by the Package to search for specific xml files) */
        const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        const string numberRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        const string CustomRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
        const string customPropRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string docNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        const String version2010 = "14.0";
        const String version2007 = "12.0";
        const String version2003 = "11.0";
        const String versionXP = "10.0";

        private const string emptyListCounter = "A";


        private String inputFilename;
        private String outputFilename;
        private String inputName;
        private String output_Pipeline;
        


        /// <summary>
        /// Docx file (zip) package resolver (to access files from the docx)
        /// </summary>
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

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrList = new ArrayList(); // not used ?

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrListLang = new ArrayList();

        /// <summary>
        /// 
        /// </summary>
        ArrayList arrCaptionProdnote = new ArrayList();

        /// <summary>
        /// 
        /// </summary>
        String strImage = "";

        /// <summary>
        /// 
        /// </summary>
        String strImageExt = "";

        /// <summary>
        /// 
        /// </summary>
        String sectionPagetype = "";

        /// <summary>
        /// 
        /// </summary>
        String getAuthor = "";
        
        /// <summary>
        /// 
        /// </summary>
        String getTitle = "";
        
        /// <summary>
        /// 
        /// </summary>
        String getYear = "";
        
        /// <summary>
        /// 
        /// </summary>
        String storeHyperId = "";
        
        /// <summary>
        /// 
        /// </summary>
        String caption, message = "";
        
        /// <summary>
        /// 
        /// </summary>
        String str;
        
        /// <summary>
        /// 
        /// </summary>
        String final;
        
        /// <summary>
        /// 
        /// </summary>
        String strImgName;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 AbbrAcrflag = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 listflag = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 listHeadingFlag = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 imgId = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 pageNum = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 flagCheck = 0;
        
        /// <summary>
        /// 
        /// </summary>
        Int16 length;
        
        /// <summary>
        /// 
        /// </summary>
        int setHyperLinkFlag = 0;
        
        /// <summary>
        /// 
        /// </summary>
        int listMasterSubFlag = 0;
        
        /// <summary>
        /// 
        /// </summary>
        int checkSectionFront = 0;
        
        /// <summary>
        /// 
        /// </summary>
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

        int rowspan = 0;
        int setToc = 0;
        int set_tabToc = 0;
        int set_Toc = 0;
        int bdoflag = 0;
        int rtlFlag = 0;

        int captionFlag = 0;
        int hyperlinkFlag = 0;
        int testRun = 0;

        int set = 0;
        int setbookmark = 0;
        int checkCverpage = 0;

        int pageId = 0;
        int flagcounter = 0;

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
        String prevHeadNumId = "";
        String baseNumId = "";
        String baseAbsId = "";
        Hashtable hTable = new Hashtable();
        Hashtable listCounters = new Hashtable();
        Hashtable headingCounters = new Hashtable();
        int objectId = 0;
        String absVal = "";
        String numFormat = "";
        String lvlText = "";
        String lStartOverride = "";
        String lStart = "";
        String headingInfo = "";
        PackageRelationship relationship = null;
        PackageRelationship imgRelationship = null;
        PackageRelationship numberRelationship = null;
        PackageRelationship customRelationship = null;
        PackageRelationship footrelationship = null;
        PackageRelationship endrelationship = null;
        Hashtable listMathMl;
        int footNoteFlag = 0;
        int sidebarFlag = 0;
        int mainFlag = 0;
        int img_Flag = 0;
        #endregion


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
        public DaisyClassSample(String inputName, String input, String output, Hashtable listMathMl, Package packInput, string output_Pipeline) {
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

        #region Exemple de "stack" utilisé par le plugin
        /// il y a quasiment une 10aine de pile gérer plus ou moins de la même façon

        const String headingZeroLvl = "0";
        const String headingOneLvl = "1";
        const String headingSixLvl = "6";

        // Voir fichier Common2.xsl:40
        /// <summary>
        /// Function to Push current document level value to the general stack
        /// (limit the level amount to 6, and seems to only push
        /// the next level from the previous one in the stack)
        /// </summary>
        /// <param name="level"></param>
        /// <returns></returns>
        public String PushLevel(String level) {
            if (stackList.Count > 0) {
                if (Convert.ToInt16(level) > 6 && stackList.Count == 6) {
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

        /* Voir 
         * - Common.xsl:308
         * - Common.xsl:1086
         * - Common.xsl:2478
         * - Common2.xsl:315
         * - Common2.xsl:755
         * - Common2.xsl:766
         * 
         */
        /// <summary>
        /// peek (retrieve without deleting) the last level of the document from the stack
        /// </summary>
        /// <returns></returns>
        public String PeekLevel() {
            if (stackList.Count > 0) {
                return stackList.Peek();
            } else {
                return headingZeroLvl;
            }
        }

        /* Voir : 
         * - Common2.xsl:320
         * - Common2.xsl:327
         */
        /// <summary>
        /// pop (retrieve and delete) the last level of the document from the stack
        /// </summary>
        /// <returns></returns>
        public String PoPLevel() {
            return stackList.Pop();
        }

        #endregion

        #region Exemples d'utilisation de compteur et variables
        
        // j'ai trouvé pas mal de fonction copier/coller sur le modèle suivant

        /* Voir : 
         * - Common3.xsl:1763
         * - Common3.xsl:1779
         * - Common3.xsl:1784
         */
        /// <summary>
        /// Increment and return as a string the pageToc counter (number of pages before the toc)
        /// </summary>
        /// <returns></returns>
        public string PageForTOC() {
            pageToc++;
            return pageToc.ToString();
        }


        /* Voir : 
         * - Common.xsl:1871
         * - Common.xsl:1877
         * - Common2.xsl:640
         * - Common2.xsl:684
         */
        // petite variation de la précédente qui modifie d'autres variables
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

        /* Voir : 
         * - Common3.xsl:1805
         * - Common3.xsl:1863
         * - Common3.xsl:1921
         * - Common3.xsl:1980
         * - Common3.xsl:2029
         */
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
        // Une version existe aussi pour les majuscules dans la classe DaisyClass


        /* Voir : 
         * - Common3.xsl:1817
         * - Common3.xsl:1875
         * - Common3.xsl:1933
         * - Common3.xsl:1994
         * - Common3.xsl:2041
         */
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
        // Je pense que la fonction est simplifiable mais a vérifier selon l'appel dans les xslt
        // Une version existe aussi pour les majuscules dans la classe DaisyClass
        #endregion

        #region Exemple d'utilisation de tableaux pour la gestion des footnotes

        /* Voir : 
         * - Common.xsl:1444
         * - Common.xsl:1455
         */
        /// <summary>
        /// Function to add Footnote to an Array
        /// </summary>
        /// <param name="inNum"></param>
        /// <returns></returns>
        public String AddFootNote(string inNum) {
            arrListNote.Add(inNum);
            return inNum;
        }

        /* Voir :
         * - Common3.xsl:131
         */
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

        #endregion

        #region Exemple de gestion de "Flag"
        // je ne sais pas pourquoi, mais ils ont préféré utilisé des entiers (cours ou non) 
        // plutot que des booleen pour gérer leur "flag"s 
        // (si a la limite ils avaient utilisé un entier et des opérateurs bitwise j'aurai compris)
        
        /* Voir : 
         * - Common.xsl:354
         * - Common.xsl:358
         * - Common.xsl:393
         * - Common.xsl:398
         * - Common.xsl:433
         * - Common.xsl:438
         * - Common.xsl:760
         * - Common.xsl:764
         * - Common.xsl:798
         * - Common.xsl:803
         * - Common.xsl:837
         * - Common.xsl:842
         */
        /// <summary>
        /// Function used to set the Abbreviations/Acronyms flag
        /// </summary>
        /// <returns></returns>
        public String SetAbbrAcrFlag() {
            AbbrAcrflag = 1;
            return "1";
        }

        /* voir:
         * - Common.xsl:470
         * - Common.xsl:481
         * - Common.xsl:905
         * - Common.xsl:914
         */
        /// <summary>
        /// Function used to reset the Abbreviations/Acronyms flag
        /// </summary>
        /// <returns></returns>
        public String ReSetAbbrAcrFlag() {
            AbbrAcrflag = 0;
            return "0";
        }

        /* Voir:
         * - Common.xsl:348
         * - Common.xsl:387
         * - Common.xsl:427
         * - Common.xsl:606
         * - Common.xsl:754
         * - Common.xsl:792
         * - Common.xsl:831
         * - Common.xsl:1072
         * - Common.xsl:1096
         * - Common2.xsl:749
         */
        /// <summary>
        /// Function used to return the Abbreviations/Acronyms flag value
        /// </summary>
        /// <returns></returns>
        public int AbbrAcrFlag() {
            return AbbrAcrflag;
        }

        #endregion

        #region fonctions utilitaires

        /* Voir:
         * - Common.xsl:2503
         * - Common.xsl:2705
         */
        /*Function used to compare two Headings info*/
        public String CompareHeading(String strA, String strB) {
            String value = "0";
            if (strA != "" && strB != "") {
                if (strA.Substring(0, strA.Length - 1).Equals(strB.Substring(0, strB.Length - 1)))
                    value = "1";
            }
            return value;

        }

        /* voir : 
         * - Common2.xsl:442
         */
        public String DiffLevel(String opLvl, String PeekLvl) {
            return Convert.ToString(Convert.ToInt16(opLvl) - Convert.ToInt16(PeekLvl));
        }

        /* voir : 
         * - Common2.xsl:474
         */
        public String ReduceOne(String opLvl) {
            return Convert.ToString(Convert.ToInt16(opLvl) - 1);
        }

        // Dans le genre fonction qui utilise des paramètres et fichiers externes
        // C'est le genre de fonction dont je ne suis pas sur que tu puisses 
        // refaire entièrement le traitement en xslt, mais il y a des trucs
        // qui aurait pu faire partie des paramètres de la feuilles xslt plutot que 
        // de cette classe ...

        /* Voir : 
        * - Common3.xsl:440
        * - Common3.xsl:443
        * - Common3.xsl:446
        * - Common3.xsl:865
        * - Common3.xsl:1184
        * - Common3.xsl:1187
        * - Common3.xsl:1190
        */
        /// <summary>
        /// - In the input document name (inputName), replace all spaces by underscore
        /// - If the output is set to the %APPDATA%/SaveAsDAISY folder and if a png image ending by the "id" parameter exists in it, copy it to the "output_pipeline" folder
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


        /* voir:
         * - Common.xsl:2771
         * - Common.xsl:2778
         * - Common.xsl:2796
         * - Common.xsl:2803
         * - Common.xsl:2853
         * - Common.xsl:2858
         * - Common.xsl:2863
         */
        /// <summary>
        /// Copy the list of counters of the base counter ID to the numID one
        /// </summary>
        /// <param name="numId">Targeted ID of list of counters</param>
        /// <returns>"1" if the numID is defined and different from the base number ID, "2" otherwise</returns>
        public string CopyToCurrCounter(String numId) {
            if (numId != baseNumId && numId != "" && numId != "0") {
                for (int i = 0; i < 9; i++) {
                    if (((ArrayList)headingCounters["List" + baseNumId])[i].ToString() != DaisyClassSample.emptyListCounter)
                        ((ArrayList)headingCounters["List" + numId])[i] = ((ArrayList)headingCounters["List" + baseNumId])[i];
                }
                return "1";
            } else
                return "2";

        }

        #endregion

        #region Exemple plus compliqué, utilisé pour la gestion des titre si j'ai bien compris
        // Dans le genre fonctions un peu plus chiadé, mais je sais pas si t'as besoin de regarder ça tout de suite XD
        // LA fonction qui est appelé dans la xslt

        /* voir:
         * - Common.xsl:2585
         * - Common.xsl:2761
         * - Common.xsl:2845
         */
        /// <summary>
        /// 
        /// </summary>
        /// <param name="currLvl">Current level number</param>
        /// <param name="location"></param>
        /// <param name="absId"></param>
        /// <returns></returns>
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

        // et les fonctions utilitaires qui trainent dans le fichier 
        // (sans documentation évidemment)
        /* voir:
         * - Common.xsl:2769
         * - Common.xsl:2776
         * - Common.xsl:2791
         * - Common.xsl:2797
         * - Common.xsl:2804
         */
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

        /* voir:
         * - Common.xsl:2770
         * - Common.xsl:2777
         * - Common.xsl:2798
         * - Common.xsl:2805
         * - Common.xsl:2852
         * - Common.xsl:2857
         * - Common.xsl:2862
         */
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


        // Une liste pour forcer le remplacement d'un numéro ? (pas sur de l'utilité)
        // Appelé StartHeadingString (pas d'appel direct dans les xslt)
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

        /* voir :
         * - Common.xsl:2772
         * - Common.xsl:2779
         * - Common.xsl:2799
         * - Common.xsl:2806
         * - Common2.xsl:36
         * Appelé aussi dans la fonction AddCurrHeadLevel 
         */
        /// <summary>
        /// 
        /// </summary>
        /// <param name="iLvl"></param>
        /// <param name="numId"></param>
        /// <param name="absId"></param>
        public void IncrementHeadingCounters(String iLvl, String numId, String absId) {
            if (numId != "" && absId != "") {
                String tempId = "";
                tempId = CheckAbstCounter(numId, absId);

                switch (((ArrayList)startHeadingItem["List" + tempId])[Convert.ToInt16(iLvl)].ToString()) {
                    case "Inc":

                        switch (((ArrayList)headingCounters["List" + numId])[Convert.ToInt16(iLvl)].ToString()) {
                            case DaisyClassSample.emptyListCounter:
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
                    ((ArrayList)headingCounters["List" + numId])[i] = DaisyClassSample.emptyListCounter;
                }
            }
        }

        // Appelé par IncrementHeadingCounters
        /// <summary>
        /// Check if a start heading item exists for the numID list and return numID it does
        /// If not, check if absID is the Base absolute ID and return the base number ID if so
        /// In other cases, returns an empty string
        /// </summary>
        /// <param name="numId"></param>
        /// <param name="absId"></param>
        /// <returns></returns>
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

        // appelé uniquement par la fonction AddCurrHeadLevel dans la classe d'origine
        /// <summary>
        /// 
        /// </summary>
        /// <param name="prevHeadLvl"></param>
        /// <param name="currLvl"></param>
        /// <param name="location"></param>
        /// <param name="absId"></param>
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

        /* Voir : 
         * - Common.xsl:2792
         * - Common.xsl:2827
         */
        /// <summary>
        /// Start a new Heading counter list if the "numId" has not already a list of counters associated
        /// If no counters are defined, base number ID and base absolute ID are defined
        /// </summary>
        /// <param name="numId">heading number ID</param>
        /// <param name="absId">absolute ID</param>
        /// <returns>"1"</returns>
        public String StartNewHeadingCounter(String numId, String absId) {
            if (headingCounters.Count == 0) {
                baseNumId = numId;
                baseAbsId = absId;
            }

            if (!headingCounters.ContainsKey("List" + numId)) {
                headingCounters.Add("List" + numId, new ArrayList(9));

                for (int i = 0; i < 9; i++)
                    ((ArrayList)headingCounters["List" + numId]).Add(DaisyClassSample.emptyListCounter);
            }

            return "1";
        }

        

        #endregion

    }
}