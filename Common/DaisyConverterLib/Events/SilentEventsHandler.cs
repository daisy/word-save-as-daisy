using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Daisy.SaveAsDAISY.Conversion.Events {

    /// <summary>
    /// Completely silent event handler with default behavior
    /// - Do not parse subdocument
    /// - Continue process on lost elements found
    /// </summary>
	public class SilentEventsHandler : IConversionEventsHandler {
        public void OnStop(string message) {
        }

        public bool AskForTranslatingSubdocuments() {
            return false;
        }

        public void OnError(string errorMessage) {
        }

        public void OnStop(string message, string title) {
            OnStop(message);
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

        public void OnUnknownError(string error) {

        }

        public void OnUnknownError(string title, string details) {

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

        public void onPreprocessingError(string inputPath, string errors) {
        }

        public void onPreprocessingSuccess() {
        }

        public bool AskForTrackConfirmation() {
            return true;
        }
    }

}
