using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.Conversion.Events {
    /// <summary>
    /// Conversion Events
    /// 
    /// More accuratly, handle interactions between the user and the converter (output and input)
    /// Maybe rename this as IConverterInteractionsHandler
    /// </summary>
    public interface IConversionEventsHandler {

        void OnStop(string message);
        bool AskForTranslatingSubdocuments();

        bool AskForTrackConfirmation();
        void OnError(string errorMessage);
        void OnStop(string message, string title);

        /// <summary>
        /// Called when conversion has finished
        /// </summary>
        void OnSuccess();

        #region Preprocess events
        void onDocumentPreprocessingStart(string inputPath);

        void onPreprocessingCancel();

        void onPreprocessingError(string inputPath, string errors);

        void onPreprocessingSuccess();
        #endregion Preprocess events

        #region Conversion events from word to dtbook XML

        /// <summary>
        /// Method called when the conversion of a list of document starts
        /// </summary>
        /// <param name="documentLists"></param>
        /// <param name="conversion"></param>
        void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion);


        /// <summary>
        /// Function called when the conversion of a document starts
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion);


        /// <summary>
        /// Method called when the conversion of a word documents list to the dtbook xml is successful (before post-processing)
        /// </summary>
        /// <param name="documentLists"></param>
        /// <param name="conversion"></param>
        void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion);

        /// <summary>
        /// Method called when the conversion of a word document to the dtbook xml is successful (before post-processing)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion);


        

        void onConversionCanceled();


        /// <summary>
        /// Progress message should indicate progression on the whole conversion process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void onProgressMessageReceived(object sender, EventArgs e);

        /// <summary>
        /// Feedback message should be informations like XSLT informations, warning and errors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">to be casted to DaisyEventArgs</param>
        void onFeedbackMessageReceived(object sender, EventArgs e);

        /// <summary>
        /// Validation feedback message (most probably validation errors)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e">to be casted to DaisyEventArgs</param>
        void onFeedbackValidationMessageReceived(object sender, EventArgs e);

        void OnSuccessMasterSubValidation(string message);

        /// <summary>
        /// Called when an unknown error was raised during conversion
        /// </summary>
        /// <param name="error"></param>
        void OnUnknownError(string error);
        void OnUnknownError(string title, string details);

        /// <summary>
        /// Called when a validation error has occured
        /// </summary>
        /// <param name="error"></param>
        /// <param name="inputFile"></param>
        /// <param name="outputFile"></param>
        void OnValidationErrors(List<ValidationError> errors, string outputFile);

        /// <summary>
        /// Called after conversion if lost elements have been fond 
        /// (elements found in a document that will not be converted, like TOC of subdocuments)
        /// </summary>
        /// <param name="outputFile"></param>
        /// <param name="elements"></param>
        void OnLostElements(string outputFile, ArrayList elements);

        /// <summary>
        /// Request user if he wants to continue the conversion with lost elements found
        /// </summary>
        /// <returns></returns>
        bool IsContinueDTBookGenerationOnLostElements();



        void OnMasterSubValidationError(string error);

        #endregion Conversion events from word to dtbook XML

        #region Post processing events
        /// <summary>
        /// Method called when post processing starts
        /// </summary>
        /// <param name="conversion"></param>
        void onPostProcessingStart(ConversionParameters conversion);
        /// <summary>
        /// Method called when the post processing pass has successfully finished 
        /// </summary>
        /// <param name="conversion"></param>
        void onPostProcessingSuccess(ConversionParameters conversion);
        #endregion Post processing events

    }



}