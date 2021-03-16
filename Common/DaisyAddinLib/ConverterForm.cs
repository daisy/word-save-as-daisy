/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * 
 */

using System;
using System.IO;
using System.Data;
using System.Text;
using System.Drawing;
using System.Threading;
using System.Resources;
using System.Collections;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using Daisy.DaisyConverter.DaisyConverterLib;
using System.Reflection;
using System.Xml.Xsl;

namespace Daisy.DaisyConverter.DaisyConverterLib
{

    public partial class ConverterForm : Form
    {
        delegate void WorkCompleteCallback(Exception e);

        private int size = 1;
        private float step;
        private Hashtable table, listMathMl;
        private Exception exception;
        private ArrayList lostElements;
        private ResourceManager manager;
        private AbstractConverter converter;
        private bool cancel, isDirect, converting, computeSize;
        private string inputFile, outputFile, validationError, output_Pipeline;
        string control = "";
        string path_For_Pipeline = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\pipeline-lite-ms";
        /* Constructor which initializes all components of the Form */
        public ConverterForm(AbstractConverter converter, string inputFile, string outputFile, ResourceManager manager, bool isDirect, Hashtable listMathMl, Hashtable table, string control, string output_Pipeline)
        {
            InitializeComponent();
            this.converter = converter;
            this.inputFile = inputFile;
            this.outputFile = outputFile;
            this.manager = manager;
            this.isDirect = isDirect;
            this.table = table;
            this.listMathMl = listMathMl;
            lblProgress.Text = "Loading translation details. Please wait for a while...";
            lostElements = new ArrayList();
            this.control = control;
            this.output_Pipeline = output_Pipeline;
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

        /* Returns bool value*/
        public bool Canceled
        {
            get
            {
                return cancel;
            }
        }

        /* Returns boolean value if there are any errors occured during translation*/
        public bool HasLostElements
        {
            get
            {
                return lostElements.Count > 0;
            }
        }

        /* Returns errors of converted XML file checking with DTD*/
        public string ValidationError
        {
            get
            {
                return validationError;
            }
        }


        /* Returns all errors occured during translation*/
        public ArrayList LostElements
        {
            get
            {
                return lostElements;
            }
        }

        /*Function  which starts DAISY Translation*/
        private void DoConvert()
        {
            try
            {
                converter.RemoveMessageListeners();
                converter.AddProgressMessageListener(new AbstractConverter.MessageListener(ProgressMessageInterceptor));
                converter.AddFeedbackMessageListener(new AbstractConverter.MessageListener(FeedbackMessageInterceptor));
                converter.AddFeedbackValidationListener(new AbstractConverter.MessageListener(FeedbackValidationInterceptor));
                converter.DirectTransform = this.isDirect;
                this.computeSize = true;
                converter.ComputeSize(this.inputFile, this.table);
                this.step = Convert.ToSingle(this.progressBar1.Maximum - this.progressBar1.Value) / this.size;
                this.computeSize = false;
                converter.Transform(this.inputFile, this.outputFile, this.table, this.listMathMl, true, output_Pipeline);
                WorkComplete(null);
            }
            catch (XsltException e) {
                AddinLogger.Error(e);
                AddinLogger.Error(e.InnerException);
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(e.InnerException.Message, "InnerException data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                WorkComplete(e);
            }
            catch (Exception e)
            {
            	AddinLogger.Error(e);
            	MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                WorkComplete(e);
            }
        }

        /* Function which Increments the value of progressbar*/
        private void ProgressMessageInterceptor(object sender, EventArgs e)
        {
            try
            {
                if (this.computeSize)
                {
                    this.size++;
                    if (this.size <= 300)
                    {
                        this.progressBar1.Increment(1);
                        float percentShow = (this.progressBar1.Value * 100) / this.progressBar1.Maximum;
                        if (Directory.Exists(path_For_Pipeline))
                        {
                            if (control == "DaisySingle" || control == "DaisyTabSingle" || control == "Daisy" || control == "Button1")
                            {
                                lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                                lblProgress.Visible = true;
                            }
                            else
                            {
                                lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                                if (percentShow == 100.00)
                                {
                                    lblProgress.Text = "Initiating daisy pipeline lite.Please wait...";
                                    lblProgress.Visible = true;
                                }
                            }
                        }
                        else
                        {
                            lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                            lblProgress.Visible = true;
                        }
                    }
                }
                else
                {
                    if (this.step < (Convert.ToSingle(this.progressBar1.Maximum - this.progressBar1.Value) / this.size) )
                        this.progressBar1.Increment(1);

                    float percentShow = (this.progressBar1.Value * 100) / this.progressBar1.Maximum;
                    if (Directory.Exists(path_For_Pipeline))
                    {
                        if (control == "DaisySingle" || control == "DaisyTabSingle" || control == "Daisy" || control == "Button1")
                        {
                            lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                            lblProgress.Visible = true;
                        }
                        else
                        {
                            lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                            if (percentShow == 100.00)
                            {
								lblProgress.Text = "Initiating daisy pipeline lite.Please wait...";
                                lblProgress.Visible = true;
                            }

                        }
                    }
                    else
                    {
                        lblProgress.Text = "Translation Process: " + percentShow.ToString() + "% " + "" + "completed";
                        lblProgress.Visible = true;
                    }
                    this.size--;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Application.DoEvents();
            if (cancel)
            {
                // As we need to leave converter.DaisyToOox, throw an exception
                throw new CancelledException();
            }
        }

        /* Function to Add  all the Messages to an Array */
        private void FeedbackMessageInterceptor(object sender, EventArgs e)
        {
            string message = ((DaisyEventArgs)e).Message;
            string messageValue = null;
            if (message.Contains("Cover Pages"))
                message = message.Replace("Cover Pages", "Cover Page");

            if (message.Contains("|"))
            {
                string[] messageKey = message.Split('|');

                int index = messageKey[0].IndexOf('%');
                // parameters substitution
                if (index > 0)
                {
                    string[] param = messageKey[0].Substring(index + 1).Split(new char[] { '%' });
                    messageValue = manager.GetString(messageKey[0].Substring(0, index));

                    if (messageValue != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            messageValue = messageValue.Replace("%" + (i + 1), param[i]);
                        }
                    }
                }
                else
                {
                    messageValue = manager.GetString(messageKey[0]);
                }

                if (messageValue != null && !lostElements.Contains(messageKey[1] + messageValue) && messageKey[1]!="")
                {
                    lostElements.Add(messageKey[1] + messageValue);
                }
            }
            else
            {
                int index = message.IndexOf('%');
                // parameters substitution
                if (index > 0)
                {
                    string[] param = message.Substring(index + 1).Split(new char[] { '%' });
                    messageValue = manager.GetString(message.Substring(0, index));

                    if (messageValue != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            messageValue = messageValue.Replace("%" + (i + 1), param[i]);
                        }
                    }
                }
                else
                {
                    messageValue = manager.GetString(message);
                }

                if (messageValue != null && !lostElements.Contains(messageValue))
                {
                    lostElements.Add(messageValue);
                }
            }


        }

        /* Function to Capture all validity errors*/
        private void FeedbackValidationInterceptor(object sender, EventArgs e)
        {
            validationError = ((DaisyEventArgs)e).Message;

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
                    string[] files = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SaveAsDAISY\");
                    foreach (string file in files)
                    {
                        if (file.Contains(".jpg") || file.Contains(".JPG") || file.Contains(".PNG") || file.Contains(".png"))
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Function to cancel the Translation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(manager.GetString("QuitConfirmation"),"SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
                cancel = true;
            else
                cancel = false;
        }

        /*Function to load the Form*/
        private void ConverterForm_Load(object sender, EventArgs e)
        {
        }

		/*Function which starts the Translation process*/
		private void ConverterForm_Shown(object sender, EventArgs e)
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
