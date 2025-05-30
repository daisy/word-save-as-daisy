using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Resources;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using MSword = Microsoft.Office.Interop.Word;
using System.Net;
using System.Xml;

namespace Daisy.SaveAsDAISY.Addins.Word2007  
{
    public partial class SuggestedReferences : Form 

    {
        MSword.Document Activedoc;
        ArrayList FtNtRef;
        string FtNtText = null;
        Object refernc ="";
        string textboxvalue = null;
        Microsoft.Office.Interop.Word.Range footnotetext;
        MSword.Range selectRng;
        MSword.Range selectRngToModify;

        public SuggestedReferences(MSword.Document doc,ArrayList FootntRefrns,String  FootNoteText)
        {
            InitializeComponent();
            Activedoc = doc;
            FtNtRef = FootntRefrns;
            textboxvalue = richTextBox1.Text = FootNoteText;

            footnotetext = Activedoc.Application.Selection.Range;

           
            DataGridViewTextBoxColumn searchresult = new DataGridViewTextBoxColumn();
            //DataGridViewTextBoxColumn linenumber = new DataGridViewTextBoxColumn();
            //DataGridViewButtonColumn select = new DataGridViewButtonColumn();
            DataGridViewCheckBoxColumn chkSelect = new DataGridViewCheckBoxColumn();

            searchresult.HeaderText = "Context";
            searchresult.Width = 350;

            //linenumber.HeaderText = "Nth Occurance from Footer";

            //select.HeaderText = "Connect";
            //select.UseColumnTextForButtonValue = true;
            //select.Text="select";

            chkSelect.HeaderText = "Select";
            if(FootntRefrns.Count > 6)
                chkSelect.Width = 50;
            else
                chkSelect.Width = 67;

            SuggstRefdataGridView.Width = 420;
            SuggstRefdataGridView.Columns.Add(searchresult);
            //SuggstRefdataGridView.Columns.Add(linenumber);
            //SuggstRefdataGridView.Columns.Add(select);
            SuggstRefdataGridView.Columns.Add(chkSelect);
            SuggstRefdataGridView.Columns[0].ReadOnly = true;

            for (int i = 0; i < FtNtRef.Count; i++)
            {
                string abc = FtNtRef[i].ToString();
                string[] pqr = abc.Split('|');
                bool firstRow = i == 0 ? true : false;
                SuggstRefdataGridView.Rows.Add(pqr[0].ToString(), firstRow);
            }

            //Shaby : Select the first row by default.
            Object StartRange;
            Object EndRange;
            selectRange(0, out StartRange, out EndRange);
            modifyRange(ref StartRange, ref EndRange);
       
        }      

        private void SuggstRefdataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //Check if the checkbox is clicked
            if (e.ColumnIndex == 1)
            {
                bool chkBoxSelected = (bool)SuggstRefdataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue;
                //Unckeck checkboxes in other rows
                if (chkBoxSelected)
                {
                    foreach (DataGridViewRow row in SuggstRefdataGridView.Rows)
                    {
                        if (row.Index != e.RowIndex)
                            row.Cells[e.ColumnIndex].Value = false;
                    }
                    button1.Enabled = true;


                    Object StartRange;
                    Object EndRange;
                    selectRange(e.RowIndex, out StartRange, out EndRange);

                    modifyRange(ref StartRange, ref EndRange);

                }
                else
                {
                    button1.Enabled = false;
                }
            }
        }

        private void modifyRange(ref Object StartRange, ref Object EndRange)
        {
            string s = Activedoc.Range(ref StartRange, ref EndRange).Text;
            //Shaby(start): Once the clue selection is done, modify the select range so as to give footnotes only to the words prior to the reference.
            //rowIndex 
            Object BodyText = Activedoc.Application.Selection.Range.Text;
            string strBodyText = BodyText.ToString().ToLower();
            EndRange = Convert.ToInt64(StartRange) + strBodyText.IndexOf(((string)refernc).ToLower()) + ((string)refernc).Length;
            selectRngToModify = Activedoc.Range(ref StartRange, ref EndRange);
            //shaby(end)
        }

        private void selectRange(int RowIndex, out Object StartRange, out Object EndRange)
        {
            //int rowIndex = SuggstRefdataGridView.Rows[e.RowIndex].Cells[1].RowIndex;
            int rowIndex = SuggstRefdataGridView.Rows[RowIndex].Cells[1].RowIndex;
            string abc1 = FtNtRef[rowIndex].ToString();
            string[] pqr1 = abc1.Split('|');
            FtNtText = pqr1[0].ToString();
            StartRange = pqr1[1];
            EndRange = pqr1[2];
            refernc = pqr1[3];
            selectRng = Activedoc.Range(ref StartRange, ref EndRange);
            selectRng.Select();
        }

        private void SuggstRefdataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                Object StartRange;
                Object EndRange;
                selectRange(e.RowIndex, out StartRange, out EndRange);
               
                #region PageNumber RnD - Commented਍ഀ
                //Object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;਍ഀ
                //object missing = System.Type.Missing;਍ഀ
                //Activedoc.Application.Selection.Fields.Add(selectRng, ref CurrentPage, ref missing, ref missing);਍ഀ
                ////Microsoft.Office.Interop.Word.WdFieldType pageNoType = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPageRef;਍ഀ
                ////int x = (int)Enum.Parse(typeof(Microsoft.Office.Interop.Word.WdFieldType), Enum.GetName(typeof(Microsoft.Office.Interop.Word.WdFieldType), pageNoType));  ਍ഀ
                #endregion਍ഀ

            }
        }

        private void SuggstRefdataGridView_CellEnter(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //Shaby (start) : below line of code to delete the previously existing reference.
                //selectRngToModify.Text = selectRngToModify.Text.ToLower().Remove(selectRngToModify.Text.IndexOf(((string)refernc).ToLower()), ((string)refernc).Length);
                //Shaby (start) : below block of code to delete the previously existing reference without loosing its formatting.
                Object referenceRangeStart = selectRngToModify.Start + selectRngToModify.Text.ToLower().IndexOf(((string)refernc).ToLower());
                Object referenceRangeEnd = (int)referenceRangeStart + ((string)refernc).Length;
                MSword.Range referenceRange = Activedoc.Range(ref referenceRangeStart, ref referenceRangeEnd);
                referenceRange.Text ="" ;
                //Shaby (end)

                Object text = textboxvalue;
                object missing = Type.Missing;
                Object reference = refernc;
                Activedoc.Footnotes.Add(selectRngToModify, ref reference, ref text);
                footnotetext.Delete(ref missing, ref missing);
                this.Close();
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void bttncancle_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}