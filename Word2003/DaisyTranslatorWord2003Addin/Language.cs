using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using MSword = Microsoft.Office.Interop.Word;

namespace Daisy.SaveAsDAISY.Addins.Word2003
{
    public partial class Language : Form
    {
        MSword.Range selectedRange;
        MSword.Document currDoc;
        ArrayList tempList, lngList;

        public Language(ArrayList lngList, MSword.Range rng, MSword.Document doc)
        {
            InitializeComponent();
            selectedRange = rng;
            currDoc = doc;
            tempList = lngList;
            this.lngList = lngList;
			this.lngList.Sort(new LanguagesComparer());

            for (int i = 2; i < this.lngList.Count; i++)
                lBx_Lngs.Items.Add(tempList[i].ToString().Replace("wd", ""));

            if (lBx_Lngs.Items.Count != 0)
                lBx_Lngs.SetSelected(0, true);


        }

        private void btn_Apply_Click(object sender, EventArgs e)
        {
            try
            {
                if (currDoc.ProtectionType == Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                {
                    selectedRange.LanguageID = (MSword.WdLanguageID)lngList[lBx_Lngs.SelectedIndex + 2];
                    currDoc.Save();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("The current document is locked for editing. Please unprotect the document.", "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

		class LanguagesComparer : IComparer
		{
			public int Compare(WdLanguageID x, WdLanguageID y)
			{
				if (x == y)
					return 0;
				if (x == WdLanguageID.wdLanguageNone)
					return -1;
				if (y == WdLanguageID.wdLanguageNone)
					return 1;
				if (x == WdLanguageID.wdNoProofing)
					return -1;
				if (y == WdLanguageID.wdNoProofing)
					return 1;
				return x.ToString().CompareTo(y.ToString());
			}

			public int Compare(object x, object y)
			{
				if (!(x is WdLanguageID) || !(y is WdLanguageID))
					return -1;
				return Compare((WdLanguageID)x, (WdLanguageID)y);
			}
		}
    }
}