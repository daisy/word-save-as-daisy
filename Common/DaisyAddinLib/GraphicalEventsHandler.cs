using Daisy.SaveAsDAISY.Conversion;
using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using MSword = Microsoft.Office.Interop.Word;

namespace Daisy.SaveAsDAISY
{
    /// <summary>
    /// Conversion events handler using graphical interface (windows dialog)
    /// </summary>
    public class GraphicalEventsHandler : IConversionEventsHandler
    {
        #region Conversion progress dialog
        public ConversionProgress ProgressDialog { get; set; }
        public Thread DialogThread { get; set; } = null;

        public void TryInitializeProgress(string message, int maximum = 1, int step = 1)
        {
            if (ProgressDialog != null && !ProgressDialog.IsDisposed)
            {
                if ((DialogThread == null || !DialogThread.IsAlive))
                {
                    DialogThread = new Thread(() =>
                    {
                        ProgressDialog.ShowDialog();
                        ProgressDialog.InitializeProgress(message, maximum, step);
                    });
                    DialogThread.Start();
                    while (!ProgressDialog.Visible)
                        ;
                }
                else
                {
                    ProgressDialog.InitializeProgress(message, maximum, step);
                }
            }
        }

        private void TryShowMessage(string message, bool isProgress = false)
        {
            if (ProgressDialog != null && !ProgressDialog.IsDisposed)
            {
                if ((DialogThread == null || !DialogThread.IsAlive))
                {
                    DialogThread = new Thread(() =>
                    {
                        ProgressDialog.ShowDialog();
                        ProgressDialog.AddMessage(message, isProgress);
                        if (isProgress)
                            ProgressDialog.Progress();
                    });
                    DialogThread.Start();
                    while (!ProgressDialog.Visible)
                        ;
                }
                else // Is already started and dialog is visible
                {
                    ProgressDialog.AddMessage(message, isProgress);
                    if (isProgress)
                        ProgressDialog.Progress();
                }
            }
        }

        private void TryClosingDialog()
        {
#if !DEBUG // keep the dialog opened for debugging
            if (DialogThread != null && DialogThread.IsAlive)
            {
                Thread.Sleep(2000);
                ProgressDialog.Close();
                DialogThread.Join();
            }
#endif
        }
#endregion

        public GraphicalEventsHandler() { }

        public void LinkToProgressDialog(
            ref ConversionProgress progressDialog,
            int maximumValue = 2
        )
        {
            this.ProgressDialog = progressDialog;
        }

        #region Preprocessing
        public void onDocumentPreprocessingStart(string inputPath)
        {
            // Intialize progress bar for preprocessing (7 steps)
            TryInitializeProgress("Preprocessing " + inputPath, 7);
        }

        public void onPreprocessingCancel()
        {
            TryShowMessage("Preprocessing canceled ");
            TryClosingDialog();
        }

        public void onPreprocessingError(string inputPath, Exception errors)
        {
            Exception e = errors;
            string message = e.Message;
            while(e.InnerException != null) {
                message += "\r\n- " + e.Message;
                e = e.InnerException;
            }
            AddinLogger.Error("Some errors were reported during preprocessing: \r\n" + message);
            TryShowMessage("An error occured in preprocessing : " + errors.Message, true);
        }

        public void onPreprocessingSuccess()
        {
            TryShowMessage("Preprocessing done", true);
        }

        public bool AskForTrackConfirmation()
        {
            return MessageBox.Show(
                    Labels.TrackConfirmation,
                    "SaveAsDAISY",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information
                ) == DialogResult.Yes;
        }

        public bool? documentMustBeRenamed(StringValidator authorizedNamePattern)
        {
            string BoxText =
                authorizedNamePattern.UnauthorizedValueMessage
                + "\r\n"
                + "\r\nDo you want to save this document under a new name ?"
                + "\r\nThe document with the original name will not be deleted."
                + "\r\n"
                + "\r\n(Click Yes to save the document under a new name and use the new one, "
                + "No to continue with the current document, "
                + "or Cancel to abort the conversion)";
            
            var result = MessageBox.Show(
                BoxText,
                "Unauthorized characters in the document filename",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning
            );
            switch (result)
            {
                case DialogResult.Yes:
                    return true;
                case DialogResult.No:
                    return false;
                case DialogResult.Cancel:
                default:
                    return null; // cancel the conversion
            }
        }

        public bool userIsRenamingDocument(ref object preprocessedObject)
        {
            object missing = Type.Missing;
            MSword.Dialog dlg = ((MSword.Document)preprocessedObject).Application.Dialogs[
                MSword.WdWordDialog.wdDialogFileSaveAs
            ];
            int saveResult = dlg.Show(ref missing);
            return saveResult == -1; // ok pressed, see https://docs.microsoft.com/fr-fr/dotnet/api/microsoft.office.interop.word.dialog.show?view=word-pia#Microsoft_Office_Interop_Word_Dialog_Show_System_Object__
        }

        public bool AskForTranslatingSubdocuments()
        {
            DialogResult dialogResult = MessageBox.Show(
                "Do you want to translate the current document along with sub documents?",
                "SaveAsDAISY",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );
            return dialogResult == DialogResult.Yes;
        }
        #endregion

        #region Conversion to dtbook
        public void onDocumentListConversionStart(
            List<DocumentProperties> documentLists,
            ConversionParameters conversion
        )
        {
            TryInitializeProgress(
                "Starting documents list conversion",
                documentLists.Count + (conversion.PipelineScript != null ? 1 : 0)
            );
        }

        public void onDocumentConversionStart(
            DocumentProperties document,
            ConversionParameters conversion
        )
        {
            TryInitializeProgress(
                "Converting document " + document.InputPath,
                conversion.PipelineScript != null ? 2 : 1
            );
        }

        public void onDocumentListConversionSuccess(
            List<DocumentProperties> documentLists,
            ConversionParameters conversion
        )
        {
            TryClosingDialog();
        }

        public void onDocumentConversionSuccess(
            DocumentProperties document,
            ConversionParameters conversion
        )
        {
            TryShowMessage(
                "Successful conversion of " + document.InputPath + " to " + document.OutputPath,
                true
            );
        }
        #endregion

        #region Post processing
        public void onPostProcessingStart(ConversionParameters conversion)
        {
            TryInitializeProgress(
                "Starting pipeline processing",
                conversion.PipelineScript.StepsCount + 1
            );

        }

        public void onPostProcessingSuccess(ConversionParameters conversion)
        {
            TryShowMessage(
                "Successfully processed or converted document, result stored in "
                    + conversion.OutputPath,
                true
            );
            TryClosingDialog();
        }

        #endregion

        public void onConversionCanceled()
        {
            TryShowMessage("Canceling conversion");
            TryClosingDialog();
        }

        public void onProgressMessageReceived(object sender, EventArgs e)
        {
            TryShowMessage(((DaisyEventArgs)e).Message, true);
        }

        public void onFeedbackMessageReceived(object sender, EventArgs e)
        {
            TryShowMessage(((DaisyEventArgs)e).Message);
        }

        public void onFeedbackValidationMessageReceived(object sender, EventArgs e)
        {
            TryShowMessage(((DaisyEventArgs)e).Message);
        }

        public void OnSuccessMasterSubValidation(string message)
        {
            MasterSubValidation infoBox = new MasterSubValidation(message, "Success");
            infoBox.ShowDialog();
        }

        public void OnConversionError(Exception errors)
        {
            Exception e = errors;
            string message = e.Message;
            while (e.InnerException != null) {
                message += "\r\n- " + e.Message;
                e = e.InnerException;
            }
            AddinLogger.Error("An error occured during conversion: \r\n" + message);
            TryShowMessage("Error during conversion : " + errors.Message, true);
        }

        public void OnValidationErrors(List<ValidationError> errors, string outputFile)
        {
            Validation validationDialog = new Validation(
                Labels.FailedLabel,
                string.Join("\r\n", errors.Select(error => error.ToString()).ToArray()),
                outputFile,
                Labels.ResourceManager
            );
            validationDialog.ShowDialog();
        }

        public void OnLostElements(string outputFile, ArrayList elements)
        {
            Fidility fidilityDialog = new Fidility(
                Labels.FeedbackLabel,
                elements,
                outputFile,
                Labels.ResourceManager
            );
            fidilityDialog.ShowDialog();
        }

        public bool IsContinueDTBookGenerationOnLostElements()
        {
            DialogResult continueDTBookGenerationResult = MessageBox.Show(
                "Do you want to create audio file",
                "SaveAsDAISY",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            return continueDTBookGenerationResult == DialogResult.Yes;
        }

        public void OnSuccess()
        {
            TryClosingDialog();
            MessageBox.Show(
                Labels.SucessLabel,
                "SaveAsDAISY - Success",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information
            );
        }

        public void OnMasterSubValidationError(string error)
        {
            MasterSubValidation infoBox = new MasterSubValidation(error, "Validation");
            infoBox.ShowDialog();
        }

        public void onPostProcessingError(Exception errors)
        {
            Exception e = errors;
            string message = e.Message;
            while (e.InnerException != null) {
                message += "\r\n- " + e.Message;
                e = e.InnerException;
            }
            AddinLogger.Error("Errors reported during conversion by DAISY Pipeline 2: \r\n" + message);
            TryShowMessage("Errors reported during conversion by DAISY Pipeline 2 : \r\n" + message, true);
        }

        public void onConversionSuccess()
        {
        }

        public void onPreprocessingWarning(string message)
        {
        }

        public void onConversionWarning(string message)
        {
        }

        public void onPostProcessingInfo(string message)
        {
        }
    }
}
