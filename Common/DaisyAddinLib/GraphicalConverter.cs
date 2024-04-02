using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

using Daisy.SaveAsDAISY.Conversion;

namespace Daisy.SaveAsDAISY.Forms {
    /// <summary>
    /// Converter extension to use graphical events handler
    /// and handle a the dialogs
    /// 
    /// </summary>
    public class GraphicalConverter : Converter {
        ConversionProgress progressDialog;

        public GraphicalConverter(
            IDocumentPreprocessor preprocessor,
            WordToDTBookXMLTransform documentConverter,
            ConversionParameters conversionParameters,
            GraphicalEventsHandler eventsHandler
        ) : base(preprocessor, documentConverter, conversionParameters, eventsHandler) {
            progressDialog = new ConversionProgress();
            progressDialog.setCancelClickListener(RequestConversionCancel);
            ((GraphicalEventsHandler)this.EventsHandler).LinkToProgressDialog(ref progressDialog);
        }

        /// <summary>
        /// Request the user to validate or complete the parameters of the conversion
        /// </summary>
        /// <param name="currentDocument">Main document to extract conversion settings</param>
        /// <returns></returns>
        public ConversionStatus requestUserParameters(DocumentParameters currentDocument) {
            // reset progress bar
            progressDialog.InitializeProgress("Waiting for user settings");
            ConversionParametersForm conversionSetter = new ConversionParametersForm(currentDocument, ConversionParameters);
            if (conversionSetter.DoTranslate() == 1) {
                // get the updated settings
                ConversionParameters = conversionSetter.UpdatedConversionParameters;
                progressDialog.BringToFront();
                progressDialog.Focus();
                return ConversionStatus.ReadyForConversion;
            }
            progressDialog.Close();
            return ConversionStatus.Canceled;
        }

        /// <summary>
        /// Request the document list and conversion settings from the user
        /// </summary>
        /// <returns></returns>
        public List<string> requestUserDocumentsList() {
            MultipleSub userRequest = new MultipleSub(ConversionParameters);
            if (userRequest.DoTranslate() == 1) {
                ConversionParameters = userRequest.UpdatedConversionParameters;
                return userRequest.GetFileNames;
            }

            return null;
        }



    }
}