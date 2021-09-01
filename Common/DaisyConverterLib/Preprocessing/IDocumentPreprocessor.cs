using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Document preprocessing requires function
    /// </summary>
    public interface IDocumentPreprocessor {

        object startPreprocessing(DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ValidateName(ref object preprocessedObject, FilenameValidator authorizedNamePattern, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus CreateWorkingCopy(ref object preprocessedObject, ref DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ProcessShapes(ref object preprocessedObject, ref DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus ProcessEquations(ref object preprocessedObject, ref DocumentParameters document, Events.IConversionEventsHandler eventsHandler = null);

        ConversionStatus endPreprocessing(ref object preprocessedObject, Events.IConversionEventsHandler eventsHandler = null);

    }
}
