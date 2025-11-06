using Daisy.SaveAsDAISY.Conversion;
using System.Collections.Generic;
namespace Daisy.SaveAsDAISY.WPF
{
    public class WPFConverter : Converter
    {
        

        public WPFConverter(
            IDocumentPreprocessor preprocessor,
            ConversionParameters conversionParameters,
            WPFEventsHandler eventsHandler
        ) : base(preprocessor, conversionParameters, eventsHandler)
        {

            eventsHandler.setCancelClickListener(RequestConversionCancel);
        }

        /// <summary>
        /// Request the user to validate or complete the parameters of the conversion
        /// </summary>
        /// <param name="currentDocument">Main document to extract conversion settings</param>
        /// <returns></returns>
        public ConversionStatus requestUserParameters(DocumentProperties currentDocument)
        {
            // reset progress bar
            //progressDialog.InitializeProgress("Waiting for user settings");
            //Conversion conversionSetter = new Conversion(currentDocument, ConversionParameters);
            //if (conversionSetter.DoTranslate() == 1) {
            //    // get the updated settings
            //    ConversionParameters = conversionSetter.UpdatedConversionParameters;
            //    progressDialog.Activate();
            //    progressDialog.Focus();
            //    return ConversionStatus.ReadyForConversion;
            //}
            ((WPFEventsHandler)this.EventsHandler).Dialog.Close();
            return ConversionStatus.Canceled;
        }

        /// <summary>
        /// Request the document list and conversion settings from the user
        /// </summary>
        /// <returns></returns>
        public List<string> requestUserDocumentsList()
        {
            //MultipleSub userRequest = new MultipleSub(ConversionParameters);
            //if (userRequest.DoTranslate() == 1) {
            //    ConversionParameters = userRequest.UpdatedConversionParameters;
            //    return userRequest.GetFileNames;
            //}

            return null;
        }



    }
}
