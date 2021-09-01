using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using MSword = Microsoft.Office.Interop.Word;

namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    public partial class progressBar : Form
    {

        private bool converting;
        MSword.WdLanguageID lngId; Int16 highValue = 0;
        MSword.Document doc;
        Hashtable detectLng;
        private Exception exception;
        int count=0;

        public progressBar(MSword.Document doc)
        {
            InitializeComponent();
            this.doc = doc;
            lblProgress.Text = "Updating language details. Please wait for a while...";
            count=doc.Paragraphs.Count;
            progressBar1.Maximum = count;
        }

        /* Returns the cancelled exception*/
        public class CancelledException : Exception
        {
        }

        /* Function Returns the exception*/
        public Exception Exception
        {
            get
            {
                return this.exception;
            }
        }

       /*Function  which starts DAISY Translation*/
        private void DoConvert()
        {
            try
            {
                for (int pGh = 1; pGh <= doc.Paragraphs.Count; pGh++)
                {
                    progressBar1.Increment(1);
                    float percentShow = (this.progressBar1.Value * 100) / count;
                    lblProgress.Text = "Updation Process: " + percentShow.ToString() + "% " + "" + "completed";
                    Application.DoEvents();
                    MSword.WdLanguageID lngId = DetectLng(doc.Paragraphs[pGh].Range);

                    try
                    {
                        doc.Paragraphs[pGh].Range.LanguageID = lngId;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
                doc.Save();
                WorkComplete(null);
            }
            catch (Exception e)
            {
                WorkComplete(e);
            }
        }


        public MSword.WdLanguageID DetectLng(MSword.Range rngPgh)
        {
            MSword.WdLanguageID lngInfo;
            detectLng = new Hashtable();
            for (int wrd = 1; wrd <= rngPgh.Words.Count; wrd++)
            {
                rngPgh.Words[wrd].LanguageDetected = false;
                rngPgh.Words[wrd].DetectLanguage();
                lngInfo = rngPgh.Words[wrd].LanguageID;
                if (detectLng.Contains(lngInfo))
                {
                    detectLng[lngInfo] = Convert.ToInt16(detectLng[lngInfo]) + 1;
                }
                else
                {
                    detectLng.Add(lngInfo, 0);
                }
            }

            highValue = 0;

            foreach (DictionaryEntry myEntry in detectLng)
            {
                String curValue = myEntry.Value.ToString();
                if (Convert.ToInt16(curValue) > highValue)
                {
                    highValue = Convert.ToInt16(curValue);
                    lngId = (MSword.WdLanguageID)myEntry.Key;
                }
            }
            return lngId;
        }


        /*Function which gives value according to exceptions*/
        private void WorkComplete(Exception e)
        {
            if (e == null)
            {
                DialogResult = DialogResult.OK;
            }
            else
            {
                if (e is CancelledException)
                {
                    DialogResult = DialogResult.Cancel;
                }
                else
                {
                    this.exception = e;
                    DialogResult = DialogResult.Abort;
                }
            }
        }

        private void progressBar_Load(object sender, EventArgs e)
        {

        }

        private void progressBar_Activated(object sender, EventArgs e)
        {
            // Launch convertion
            if (!converting)
            {
                converting = true;
                Application.DoEvents();
                DoConvert();
                converting = false;
            }
        }




    }
}