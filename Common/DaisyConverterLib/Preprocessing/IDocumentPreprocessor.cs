
namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Document preprocessing requires function
    /// </summary>
    public interface IDocumentPreprocessor {
        /// <summary>
        /// Read document data from a document object.
        /// </summary>
        /// <param name="documentObject"></param>
        /// <returns></returns>
        DocumentProperties loadDocumentParameters(ref object documentObject);

        /// <summary>
        /// Reopen
        /// </summary>
        /// <param name="document"></param>
        /// <param name="eventsHandler"></param>
        /// <returns></returns>
        object startPreprocessing(DocumentProperties document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ValidateName(ref object preprocessedObject, StringValidator authorizedNamePattern, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus CreateWorkingCopy(ref object preprocessedObject, ref DocumentProperties document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ProcessShapes(ref object preprocessedObject, ref DocumentProperties document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ProcessEquations(ref object preprocessedObject, ref DocumentProperties document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus endPreprocessing(ref object preprocessedObject, Events.IConversionEventsHandler eventsHandler = null);

        void updateDocumentMetadata(ref object documentObject, DocumentProperties data);
    }
}
