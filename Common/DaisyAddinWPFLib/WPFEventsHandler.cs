using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using static Daisy.SaveAsDAISY.WPF.ConversionProgress;
using MSWord = Microsoft.Office.Interop.Word;
namespace Daisy.SaveAsDAISY.WPF
{
    public class WPFEventsHandler : Daisy.SaveAsDAISY.Conversion.Events.IConversionEventsHandler
    {

        public ConversionProgress Dialog = null;

        #region Conversion progress dialog
        public void TryInitializeProgress(string message, int maximum = 1, int step = 1)
        {
            if(Dialog == null) {
                Dialog = new ConversionProgress();
                Dialog.Closed += Dialog_Closed;
                if(cancelButtonClicked != null) {
                    Dialog.setCancelClickListener(cancelButtonClicked);
                }
                Dialog.Show();
            }
            Dialog.Dispatcher.Invoke(() => Dialog.InitializeProgress(message, maximum, step));
        }

        private event CancelClickListener cancelButtonClicked = null;

        public void setCancelClickListener(CancelClickListener cancelAction)
        {
            cancelButtonClicked = cancelAction;
            if (Dialog != null) {
                Dialog.setCancelClickListener(cancelAction);
            }
        }

        private void Dialog_Closed(object sender, EventArgs e)
        {
            Dialog = null;
        }

        private void TryShowMessage(string message, bool isProgress = false)
        {
            if (Dialog == null) {
                Dialog = new ConversionProgress();
                Dialog.Closed += Dialog_Closed;
                Dialog.Show();
            }
            Dialog.Dispatcher.Invoke(() => Dialog.AddMessage(message, isProgress));
        }



        private void TryClosingDialog(int sleepBefore)
        {
            if (Dialog == null) {
                return;
            }
            Dialog.Dispatcher.Invoke(() => {
                Thread.Sleep(sleepBefore);
                Dialog.Close();
                Dialog = null;
            });
        }
        #endregion


        #region Preprocessing
        public void onDocumentPreprocessingStart(string inputPath)
        {
            // Intialize progress bar for preprocessing / document analysis before conversion (4 steps)
            TryInitializeProgress("Preprocessing " + inputPath, 4);
        }

        public void onPreprocessingCancel()
        {
            TryShowMessage("Preprocessing canceled ");
            TryClosingDialog(3000);
        }

        public void onPreprocessingError(string inputPath, Exception errors)
        {
            Exception e = errors;
            string message = e.Message;
            while (e.InnerException != null) {
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
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information
                ) == MessageBoxResult.Yes;
        }

        public bool? documentMustBeRenamed(StringValidator authorizedNamePattern)
        {
            string BoxText =
                authorizedNamePattern.UnauthorizedValueMessage
                + "\r\n"
                + "\r\nDo you want to save this _document under a new name ?"
                + "\r\nThe _document with the original name will not be deleted."
                + "\r\n"
                + "\r\n(Click Yes to save the _document under a new name and use the new one, "
                + "No to continue with the current _document, "
                + "or Cancel to abort the conversion)";
            var result = MessageBox.Show(
                BoxText,
                "Unauthorized characters in the _document filename",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Warning
            );
            switch(result) {
                case MessageBoxResult.Yes:
                    return true; // user wants to rename the document
                case MessageBoxResult.No:
                    return false; // user wants to continue with the current document
                case MessageBoxResult.Cancel:
                    default:
                    return null; // user wants to cancel the conversion
            }
        }

        public bool userIsRenamingDocument(ref object preprocessedObject)
        {
            object missing = Type.Missing;
            MSWord.Dialog dlg = ((MSWord.Document)preprocessedObject).Application.Dialogs[
                MSWord.WdWordDialog.wdDialogFileSaveAs
            ];
            int saveResult = dlg.Show(ref missing);
            return saveResult == -1; // ok pressed, see https://docs.microsoft.com/fr-fr/dotnet/api/microsoft.office.interop.word.dialog.show?view=word-pia#Microsoft_Office_Interop_Word_Dialog_Show_System_Object__
        }

        public bool AskForTranslatingSubdocuments()
        {
            var dialogResult = MessageBox.Show(
                "Do you want to translate the current _document along with sub documents?",
                "SaveAsDAISY",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information
            );
            return dialogResult == MessageBoxResult.Yes;
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
                "Converting _document " + document.InputPath,
                conversion.PipelineScript != null ? 2 : 1
            );
        }

        public void onDocumentListConversionSuccess(
            List<DocumentProperties> documentLists,
            ConversionParameters conversion
        )
        {
            TryClosingDialog(3000);
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
                "Successfully processed or converted _document, result stored in "
                    + conversion.OutputPath,
                true
            );
            TryClosingDialog(3000);
        }

        #endregion

        public void onConversionCanceled()
        {
            TryShowMessage("Canceling conversion");
            TryClosingDialog(3000);
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
            //MasterSubValidation infoBox = new MasterSubValidation(message, "Success");
            //infoBox.ShowDialog();
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

        public void OnValidationErrors(List<SaveAsDAISY.Conversion.ValidationError> errors, string outputFile)
        {
            SaveAsDAISY.Conversion.Validation validationDialog = new SaveAsDAISY.Conversion.Validation(
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
            var continueDTBookGenerationResult = MessageBox.Show(
                "Do you want to create audio file",
                "SaveAsDAISY",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question
            );
            return continueDTBookGenerationResult == MessageBoxResult.Yes;
        }

        public void OnSuccess()
        {
            TryClosingDialog(3000);
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
            TryShowMessage("Successfull conversion", false);
            TryClosingDialog(3000);
        }

        public void onPreprocessingWarning(string message)
        {
            TryShowMessage(message, false);
        }

        public void onConversionWarning(string message)
        {
            TryShowMessage(message, false);
        }

        public void onPostProcessingInfo(string message)
        {
            TryShowMessage(message, false);
        }

    }
}
