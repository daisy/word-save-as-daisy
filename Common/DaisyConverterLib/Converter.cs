using Daisy.SaveAsDAISY.Conversion.Events;
using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Pipeline2.Scripts;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Resources.ResXFileRef;

namespace Daisy.SaveAsDAISY.Conversion
{

    /// <summary>
    /// Base class to preprocess and convert one or more preprocessed document to a single DTBook XML
    /// and then possibly to other format using the Daisy pipeline
	/// 
	/// Its more a "Document processing workflow" 
    /// </summary>
    public class Converter
    {

        private ConversionParameters conversionParameters;

        private ChainResourceManager resourceManager;

        private string validationErrorMsg = string.Empty;

        private CancellationTokenSource xmlConversionCancel;

        private ConversionStatus currentStatus = ConversionStatus.None;

        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager
        {
            set { this.ResourceManager.Add(value); }
        }


        /// <summary>
        /// Retrieve the label associated to the specified key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public string GetString(string key)
        {
            return this.ResourceManager.GetString(key);
        }

        /// <summary>
        /// Resource manager
        /// </summary>
        public System.Resources.ResourceManager ResManager
        {
            get { return this.ResourceManager; }
        }
        public ConversionParameters ConversionParameters { get => conversionParameters; set => conversionParameters = value; }
        protected ChainResourceManager ResourceManager { get => resourceManager; set => resourceManager = value; }
        protected string ValidationErrorMsg { get => validationErrorMsg; set => validationErrorMsg = value; }
        public ConversionStatus CurrentStatus { get => currentStatus; set => currentStatus = value; }
        protected IConversionEventsHandler EventsHandler { get => eventsHandler; set => eventsHandler = value; }
        protected IDocumentPreprocessor DocumentPreprocessor { get => documentPreprocessor; set => documentPreprocessor = value; }

        //public ConversionParameters Conversion { get => conversion; set => conversion = value; }

        /// <summary>
        /// Events handler class
        /// </summary>
        private IConversionEventsHandler eventsHandler;

        /// <summary>
        /// Document preprocessor, that depends on word interop version
        /// </summary>
        private IDocumentPreprocessor documentPreprocessor;

        public Converter(
            IDocumentPreprocessor preprocessor,
            ConversionParameters conversionParameters,
            IConversionEventsHandler eventsHandler = null)
        {
            this.DocumentPreprocessor = preprocessor;
            this.ConversionParameters = conversionParameters;
            this.EventsHandler = eventsHandler ?? new SilentEventsHandler();
           
            this.ResourceManager = new ChainResourceManager();
            // Add a default resource managers (for common labels)
            this.ResourceManager.Add(
                new System.Resources.ResourceManager(
                    "DaisyAddinLib.resources.Labels",
                    Assembly.GetExecutingAssembly()
                )
            );
        }

        /// <summary>
        /// Second preprocessing path to update the document metadata and extract shapess
        /// </summary>
        /// <param name="docprops"></param>
        public void PrepareForConversion(ref DocumentProperties docprops)
        {
            try {
                EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Preparing the conversion"));
                // Reopen or focus the document
                object preprocessedObject = DocumentPreprocessor.startPreprocessing(docprops, EventsHandler);
                // Update the working copy
                EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Updating the working copy ..."));
                DocumentPreprocessor.CreateWorkingCopy(ref preprocessedObject, ref docprops, EventsHandler);
                // Export shapes and Equations
                EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Processing shapes..."));
                CurrentStatus = DocumentPreprocessor.ProcessShapes(ref preprocessedObject, ref docprops, EventsHandler);
                EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Processing Math..."));
                CurrentStatus = DocumentPreprocessor.ProcessEquations(ref preprocessedObject, ref docprops, EventsHandler);
                // Close the document to let it be accessible for conversion
                CurrentStatus = DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
            }
            catch (Exception e) {
                Exception fault = new Exception("An error occured after " + CurrentStatus.ToString() + " conversion pass:\r\n" + e.Message, e);
                throw fault;
            }
        }

        /// <summary>
        /// Preprocess a given document :
        /// 
        /// </summary>
        /// <param name="inputPath"></param>
        /// <param name="resourceId"></param>
        /// <returns>A document ready for conversion or null if an error has occured or if the preprocess has been canceled</returns>
        public DocumentProperties AnalyzeDocument(string inputPath, string resourceId = null)
        {
            EventsHandler.onDocumentPreprocessingStart(inputPath);
            DocumentProperties result = new DocumentProperties(inputPath)
            {
                ResourceId = resourceId ?? null,
                ShowInputDocumentInWord = ConversionParameters.Visible
            };
            // dot not make visible subdocuments (documents with resource Id assigned)
            // Note : allow preprocessing to be "silent"
            object preprocessedObject = DocumentPreprocessor.startPreprocessing(result, EventsHandler);

            try {
                do {
                    switch (CurrentStatus) {
                        case ConversionStatus.None: // Starting by validating file name
                            EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Validating file name"));
                            CurrentStatus = DocumentPreprocessor.ValidateName(ref preprocessedObject, ConversionParameters.NameValidator, EventsHandler);
                            break;
                        case ConversionStatus.ValidatedName: // make the working copy
                            EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Name validated, creating a working copy"));
                            CurrentStatus = DocumentPreprocessor.CreateWorkingCopy(ref preprocessedObject, ref result, EventsHandler);
                            break;
                        case ConversionStatus.CreatedWorkingCopy: // Load properties and languages from the working copy
                            EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Working copy created, reading properties and languages from it"));
                            CurrentStatus = DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
                            result.updatePropertiesFromCopy();
                            break;
                        //case ConversionStatus.ProcessedShapes: // start processing math
                        //    EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Shapes processed, processing mathml"));
                        //    CurrentStatus = DocumentPreprocessor.ProcessEquations(ref preprocessedObject, ref result, EventsHandler);
                        //    break;
                        //case ConversionStatus.ProcessedMathML: // finalize preprocessing
                        //    EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("MathML processed"));
                        //    CurrentStatus = DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
                        //    break;
                    }
                } while (CurrentStatus != ConversionStatus.Canceled &&
                        CurrentStatus != ConversionStatus.Error &&
                        CurrentStatus != ConversionStatus.PreprocessingSucceeded);
            }
            catch (Exception e)
            {
                Exception fault = new Exception("An error occured after " + CurrentStatus.ToString() + " status preprossing:\r\n" + e.Message, e);
                EventsHandler.onPreprocessingError(inputPath, fault);
                throw fault;
            }
            finally
            {
                if (CurrentStatus != ConversionStatus.PreprocessingSucceeded)
                {
                    DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
                }
            }
            if (CurrentStatus != ConversionStatus.PreprocessingSucceeded)
            {
                return null;
            }
            // Check for revisions
            if (result.HasRevisions)
            {
                result.AcceptRevisions = EventsHandler.AskForTrackConfirmation();
                // To be removed later, after moving TrackChanges eval from conversion to document param object
                ConversionParameters.TrackChanges = result.AcceptRevisions ? "Yes" : "No";
            }
            else
            {
                ConversionParameters.TrackChanges = "NoTrack";
            }

            // For now, subdocuments are merged into the master document
            // if the user confirmed he wants them converted in preprocessing
            ConversionParameters.ParseSubDocuments = false;
            CurrentStatus = ConversionStatus.PreprocessingSucceeded;
            EventsHandler.onPreprocessingSuccess();
            return result;

        }

        /// <summary>
        /// Convert a single document to XML using only DAISY Pipeline 2 scripts
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        public ConversionResult ConvertWithPipeline2(DocumentProperties document, ConversionParameters conversion = null)
        {
            ConversionParameters _conversion = conversion ?? this.ConversionParameters;
            this.EventsHandler.onDocumentConversionStart(document, _conversion);
            xmlConversionCancel = new CancellationTokenSource();
            CurrentStatus = ConversionStatus.HasStartedConversion;
            
            PrepareForConversion(ref document);

            // If conversion is to be post processed output the xsl transfo to temp
            string outputDirectory = _conversion.OutputPath.EndsWith(".xml") ?
                        Directory.GetParent(_conversion.OutputPath).FullName :
                        _conversion.OutputPath;
            // For script execution, replace this by a temp directory
            //if (_conversion.PipelineScript != null) {
            //    outputDirectory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;
            //}

            // Rebuild and sanitize file name
            string outputFilename = (
                _conversion.OutputPath.EndsWith(".xml") ?
                    Path.GetFileName(_conversion.OutputPath) :
                    Path.GetFileNameWithoutExtension(document.InputPath) + ".xml"
                );

            string sanitizedName = outputFilename.Replace(",", "_");/*.Replace(" ", "_");*/

            // Rebuild output path
            document.OutputPath = Path.Combine(outputDirectory, sanitizedName);

            if (document.SubDocumentsToConvert.Count > 0 && _conversion.ParseSubDocuments) {
                string message = "Subdocuments conversion is undergoing changes and is not available for now.\r\n " +
                    "Please unlink your subdocuments to merge them into the master document before conversion.\r\n" +
                    "(Open the outline view, display and expand subdocuments, select each subdocument and click on the \"Unlink\" button)" +
                    "";
                MessageBox.Show(message, "Subdocuments unimplemented", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.EventsHandler.OnConversionError(
                    new Exception(message)
                );
                this.EventsHandler.onConversionCanceled();
                return ConversionResult.Cancel();
                // TODO : either rework script ExecuteScript method to allow multiple input, to possibly merge them
                // or rework the xsl to include conversion of other documents ?
                //List<DocumentParameters> flattenList = new List<DocumentParameters> {
                //    document
                //};
                //foreach (DocumentParameters subDocument in document.SubDocumentsToConvert) {
                //    flattenList.Add(subDocument);
                //}
                //// switch to list of document conversion
                //return this.ConvertWithPipeline2AndMerge(flattenList);
            } else {
                if (CurrentStatus == ConversionStatus.Canceled) { // Conversion is aborted
                    this.EventsHandler.onConversionCanceled();
                    return ConversionResult.Cancel();
                }

                if (_conversion.PipelineScript == null) throw new Exception("No script selected for conversion");
                // launch the pipeline post processing
                this.EventsHandler.onPostProcessingStart(_conversion);
                try {
                    _conversion.PipelineScript.ExecuteScript(document.CopyPath);
                }
                catch (JobException je) {
                    // Job finished in error, not sure if i should  return a failed result
                    // or throw back to allow a report
                    this.EventsHandler.onPostProcessingError(
                        je
                    );
                    return ConversionResult.Failed(je.Message);
                }
                catch (Exception e) {
                    CurrentStatus = ConversionStatus.Error;
                    Exception fault = new Exception("Error while converting with DAISY Pipeline 2", e);
                    this.EventsHandler.onPostProcessingError(
                        fault
                    );
                    throw fault;
                }
                this.EventsHandler.onPostProcessingSuccess(ConversionParameters);
                
                return ConversionResult.Success();
            }
            
        }

        public ConversionResult ConvertWithPipeline2AndMerge(List<DocumentProperties> documentLists)
        {
            
            if (documentLists is null) {
                throw new ArgumentNullException(nameof(documentLists));
            }
            if(ConversionParameters.PipelineScript == null) {
                throw new Exception("No script selected for conversion");
            }
            this.EventsHandler.onDocumentListConversionStart(documentLists, ConversionParameters);
            CurrentStatus = ConversionStatus.HasStartedConversion;
            string errors = "";

            try {
                XmlDocument mergeResult = null;
                if (documentLists.Count == 1) {
                    return this.ConvertWithPipeline2(documentLists[0]);
                } else {
                    string outputDirectory = ConversionParameters.OutputPath.EndsWith(".xml") ?
                        Directory.GetParent(ConversionParameters.OutputPath).FullName :
                        ConversionParameters.OutputPath;
                    // Rebuild and sanitize file name
                    string outputFilename = (
                        ConversionParameters.OutputPath.EndsWith(".xml") ?
                            Path.GetFileName(ConversionParameters.OutputPath) :
                            Path.GetFileNameWithoutExtension(documentLists[0].InputPath) + ".xml"
                        ).Replace(" ", "_").Replace(",", "_");

                    // Rebuild output path based on first document 
                    ConversionParameters.OutputPath = Path.Combine(outputDirectory, outputFilename);
                    WordToDtbook wordToDtbook = new WordToDtbook(this.EventsHandler);
                    throw new NotImplementedException("Functionnality is being changed");
                    //foreach (DocumentParameters document in documentLists) {
                    //    //document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
                    //    // use memory temp for output
                    //    //document.OutputPath = Path.GetTempFileName();
                    //    // TODO
                    //    // - convert with pipeline 2
                    //    // - Merge the conversion of the document in the final result

                    //    if(mergeResult == null) {

                    //    } else {
                    //        //finalResult = DocumentConverter.MergeDocumentInTarget(document, finalResult);
                    //    }

                    //    if (CurrentStatus == ConversionStatus.Canceled) {
                    //        this.EventsHandler.onConversionCanceled();
                    //        return ConversionResult.Cancel();
                    //    }
                    //    this.EventsHandler.onDocumentConversionSuccess(document, ConversionParameters);
                    //}
                    //DocumentConverter.finalizeAndSaveMergedDocument(mergeResult, ConversionParameters);
                }
            }
            catch (Exception e) {
                // Propagate unhandled exception
                throw new Exception(ResourceManager.GetString("TranslationFailed") + "\n"
                    + ResourceManager.GetString("WellDaisyFormat") + "\n" + " \""
                    + errors + "\n" + "Crictical issue:" + "\n" + e.Message + "\n", e);

            }
            //if (DocumentConverter.ValidationErrors.Count > 0) {
            //    this.EventsHandler.OnValidationErrors(DocumentConverter.ValidationErrors, ConversionParameters.OutputPath);
            //    return ConversionResult.FailedOnValidation(
            //        string.Join(
            //            "\r\n",
            //            DocumentConverter.ValidationErrors.Select(
            //                error => error.ToString()
            //            ).ToArray()
            //        )
            //    );
            //} else if (CurrentStatus == ConversionStatus.Canceled) {
            //    return ConversionResult.Cancel();
            //} else {
            //    this.EventsHandler.onDocumentListConversionSuccess(documentLists, ConversionParameters);

            //    if (DocumentConverter.LostElements.Count > 0) {
            //        ArrayList unconvertedElements = new ArrayList();
            //        foreach (KeyValuePair<string, List<string>> lostElementForFile in DocumentConverter.LostElements) {
            //            if (lostElementForFile.Value.Count > 0) {
            //                string lostElements = lostElementForFile.Key + ":\r\n";
            //                foreach (string lostElement in lostElementForFile.Value) {
            //                    lostElements += " - " + lostElement + "\r\n";
            //                }
            //                unconvertedElements.Add(lostElements);
            //            }
            //        }
            //        this.EventsHandler.OnLostElements(ConversionParameters.OutputPath, unconvertedElements);
            //    }
                
            //    // If post processing is requested (and can be applied even if lost elements are found)
            //    // Launch the post processing pipeline sript 
            //    // (cleaning or converting DTBook to another format Like a DAISY book)
            //    this.EventsHandler.onPostProcessingStart(ConversionParameters);
            //    try {
            //        ConversionParameters.PostProcessor.ExecuteScript(ConversionParameters.OutputPath);
            //    }
            //    catch (JobException je) {
            //        // Job finished in error, not sure if i should  return a failed result
            //        // or throw back to allow a report
            //        this.EventsHandler.onPostProcessingError(
            //            je
            //        );
            //        return ConversionResult.Failed(je.Message);
            //    }
            //    catch (Exception e) {
            //        Exception fault = new Exception("An execution error occured while running conversion in DAISY Pipeline 2:" + e.Message, e);
            //        this.EventsHandler.onPostProcessingError(
            //            fault
            //        );
            //        throw fault;
            //        //return ConversionResult.Failed(e.Message);
            //    }
            //    this.EventsHandler.onPostProcessingSuccess(ConversionParameters);
                
            //}
            //return ConversionResult.Success();

        }


        /// <summary>
        /// 
        /// </summary>
        protected void RequestConversionCancel()
        {
            xmlConversionCancel?.Cancel();
            CurrentStatus = ConversionStatus.Canceled;

        }


    }

}
