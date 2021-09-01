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
            progressDialog.setCancelClickListener(requestConversionCancel);
            ((GraphicalEventsHandler)this.eventsHandler).LinkToProgressDialog(ref progressDialog);
        }

        /// <summary>
        /// Request the user to validate or complete the parameters of the conversion
        /// </summary>
        /// <param name="currentDocument">Main document to extract conversion settings</param>
        /// <returns></returns>
        public ConversionStatus requestUserParameters(DocumentParameters currentDocument) {

            ConversionParametersForm conversionSetter = new ConversionParametersForm(currentDocument, conversion);
            if (conversionSetter.DoTranslate() == 1) {
                // get the updated settings
                conversion = conversionSetter.UpdatedConversionParameters;
                return ConversionStatus.ReadyForConversion;
            }
            return ConversionStatus.Canceled;
        }

        /// <summary>
        /// Request the document list and conversion settings from the user
        /// </summary>
        /// <returns></returns>
        public List<string> requestUserDocumentsList() {
            MultipleSub userRequest = new MultipleSub(conversion);
            if (userRequest.DoTranslate() == 1) {
                conversion = userRequest.UpdatedConversionParameters;
                return userRequest.GetFileNames;
            }

            return null;
        }



    }
}