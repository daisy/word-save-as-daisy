using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Collections;
using MSword = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Resources;
using System.Reflection;

namespace Daisy.SaveAsDAISY.Addins.Word2007
{
    public partial class frmValidate2007 : Form
    {
        ArrayList error = new ArrayList();
        MSword.Document pdoc;
        /*Default constructor*/
        public frmValidate2007()
        {
            InitializeComponent();
        }
        /*Function which displays validation errors in a datagrid*/
       /*Function which displays validation errors in a datagrid*/
        public void SetApp(ArrayList error, MSword.Document doc)
        {
            pdoc = doc;
            pdoc.Save();
            this.error = error;
            //Looping for number of errors in the error arraylist
            for (int i = 0; i < this.error.Count; i++)
            {
                string validationMsg = this.error[i].ToString();
                string[] messageKey = validationMsg.Split('|');
                string[] errMessage = messageKey[0].Split(':');
                this.dataGrid_Validation.Rows.Add(errMessage[1], "More Info");
            }
        }
        /*Function which select the location of the error on single click*/
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //if second column is selected
            if (e.ColumnIndex == 1)
            {
                //Getting error message from first column
                String str = this.dataGrid_Validation.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                for (int i = 0; i < error.Count; i++)
                {
                    string[] index = error[i].ToString().Split(':');
                    string[] messageKey = index[1].Split('|');
                    if (str == messageKey[0].ToString())
                    {
                        //showing help file
                        Help.ShowHelp(dataGrid_Validation, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.KeywordIndex, index[0]);
                    }
                }
                
            }
            //if first column is selected
            else if (e.ColumnIndex == 0)
            {
                string[] messageKey = this.error[e.RowIndex].ToString().Split('|');
                foreach (object item in pdoc.Bookmarks)
                {
                    //checking for bookmark
                    if (((MSword.Bookmark)item).Name == messageKey[1])
                    {
                        //selecting the location of the bookmark
                        ((MSword.Bookmark)item).Select();
                        //activating the word document
                        pdoc.Application.Activate();
                    }
                }
            }
        }
        /*Function which select the location of the error on double click */
        void dataGrid_Validation_CellContentDoubleClick(object sender, System.Windows.Forms.DataGridViewCellEventArgs e)
        {
            //if second column is selected
            if (e.ColumnIndex == 1)
            {
                //Getting error message from first column
                String str = this.dataGrid_Validation.Rows[e.RowIndex].Cells[e.ColumnIndex - 1].Value.ToString();
                for (int i = 0; i < error.Count; i++)
                {

                    string[] index = error[i].ToString().Split(':');
                    string[] messageKey = index[1].Split('|');
                    if (str == messageKey[0].ToString())
                    {

                        Help.ShowHelp(dataGrid_Validation, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.KeywordIndex, index[0]);
                    }
                }
                //showing help file
            }
            //if first column is selected
            else if (e.ColumnIndex == 0)
            {
                string[] messageKey = this.error[e.RowIndex].ToString().Split('|');
                //checking for bookmark
                foreach (object item in pdoc.Bookmarks)
                {
                    if (((MSword.Bookmark)item).Name == messageKey[1])
                    {
                        //selecting the location of the bookmark
                        ((MSword.Bookmark)item).Select();
                        //activating the word document
                        pdoc.Application.Activate();
                    }
                }
            }
        }
        /*Function which select the location of the error on key press */
        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if enter key is pressed
            if (e.KeyChar.Equals('\r'))
            {
                DataGridViewCell cell = dataGrid_Validation.CurrentCell;
                //if second column is selected
                if (this.dataGrid_Validation.CurrentCell.ColumnIndex == 1)
                {
                    //Getting error message from first column
                    String str = this.dataGrid_Validation.Rows[cell.RowIndex].Cells[cell.ColumnIndex - 1].Value.ToString();
                    for (int i = 0; i < error.Count; i++)
                    {
                        string[] index = error[i].ToString().Split(':');
                        string[] messageKey = index[1].Split('|');
                        if (str == messageKey[0].ToString())
                        {
                            Help.ShowHelp(dataGrid_Validation, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + "Help.chm", HelpNavigator.KeywordIndex, index[0]);
                        }
                    }
                    //showing help file
                }
                //if first column is selected
                else if (this.dataGrid_Validation.CurrentCell.ColumnIndex == 0)
                {

                    string[] messageKey = this.error[cell.RowIndex - 1].ToString().Split('|');
                    foreach (object item in pdoc.Bookmarks)
                    {
                        //checking for bookmark
                        if (((MSword.Bookmark)item).Name == messageKey[1])
                        {
                            //selecting the location of the bookmark
                            ((MSword.Bookmark)item).Select();
                            //activating the word document
                            pdoc.Application.Activate();

                        }
                    }
                }
            }
        }
        /*Function which deletes the book mark on click of close button*/
        private void btnValidate_Click(object sender, EventArgs e)
        {
            if (this.error.Count > 0)
            {
                foreach (object item in pdoc.Bookmarks)
                {
                    //Checking for Bookmark with Rule
                    if (((MSword.Bookmark)item).Name.StartsWith("Rule", StringComparison.CurrentCulture))
                    {
                        object val = ((MSword.Bookmark)item).Name;
                        //Deleting bookmark
                        pdoc.Bookmarks.get_Item(ref val).Delete();
                    }
                }
                //closing the form
                this.Close();
                //saving the word document
                pdoc.Save();
            }
        }

       
    }
}