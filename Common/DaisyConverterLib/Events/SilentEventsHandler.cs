using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Daisy.SaveAsDAISY.Conversion.Events {

    /// <summary>
    /// Completely silent event handler with default behavior
    /// - Do not parse subdocument
    /// - Continue process on lost elements found
    /// </summary>
	public class SilentEventsHandler : IConversionEventsHandler {
       
        public bool AskForTranslatingSubdocuments() {
            return false;
        }


        public void onDocumentListConversionStart(List<DocumentParameters> documentLists, ConversionParameters conversion) {

        }

        public void onDocumentConversionStart(DocumentParameters document, ConversionParameters conversion) {

        }

        public void onPostProcessingStart(ConversionParameters conversion) {

        }

        public void onDocumentListConversionSuccess(List<DocumentParameters> documentLists, ConversionParameters conversion) {

        }

        public void onDocumentConversionSuccess(DocumentParameters document, ConversionParameters conversion) {

        }

        public void onPostProcessingSuccess(ConversionParameters conversion) {

        }

        public void onConversionCanceled() {

        }

        public void onProgressMessageReceived(object sender, EventArgs e) {

        }

        public void onFeedbackMessageReceived(object sender, EventArgs e) {

        }

        public void onFeedbackValidationMessageReceived(object sender, EventArgs e) {

        }

        public void OnSuccessMasterSubValidation(string message) {

        }


        public void OnValidationErrors(List<ValidationError> errors, string outputFile) {

        }

        public void OnLostElements(string outputFile, ArrayList elements) {

        }

        public bool IsContinueDTBookGenerationOnLostElements() {
            return true;
        }

        public void OnSuccess() {

        }

        public void OnMasterSubValidationError(string error) {

        }

        public void onDocumentPreprocessingStart(string inputPath) {
        }

        public void onPreprocessingCancel() {
        }

        public void onPreprocessingSuccess() {
        }

        public bool AskForTrackConfirmation() {
            return true;
        }

        public DialogResult documentMustBeRenamed(StringValidator authorizedNamePattern) {
            // renamed automatically for silent mode, see the interface reference
            return DialogResult.No;
        }

        public bool userIsRenamingDocument(ref object preprocessedObject) {
            // Should not be used in silent mode
            throw new NotImplementedException();
        }

        public void onPreprocessingError(string inputPath, Exception errors)
        {
        }

        public void OnConversionError(Exception error)
        {
        }

        public void onPostProcessingError(Exception error)
        {
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
