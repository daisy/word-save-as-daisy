﻿using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {

    /// <summary>
    /// Base class to preprocess and convert one or more preprocessed document to a single DTBook XML
    /// and then possibly to other format using the Daisy pipeline
	/// 
	/// Its more a "Document processing workflow" 
    /// </summary>
    public class Converter {

        /// <summary>
        /// 
        /// </summary>
        private WordToDTBookXMLTransform documentConverter;
        private ConversionParameters conversionParameters;

        private ChainResourceManager resourceManager;

        private string validationErrorMsg = string.Empty;


		private Task<XmlDocument> conversionTask;
		private CancellationTokenSource xmlConversionCancel;



        private ConversionStatus currentStatus = ConversionStatus.None;

        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager {
			set { this.ResourceManager.Add(value); }
		}


		/// <summary>
		/// Retrieve the label associated to the specified key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public string GetString(string key) {
			return this.ResourceManager.GetString(key);
		}

		/// <summary>
		/// Resource manager
		/// </summary>
		public System.Resources.ResourceManager ResManager {
			get { return this.ResourceManager; }
		}

        protected WordToDTBookXMLTransform DocumentConverter { get => documentConverter; set => documentConverter = value; }
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

        public Converter(IDocumentPreprocessor preprocessor, WordToDTBookXMLTransform documentConverter, ConversionParameters conversionParameters, IConversionEventsHandler eventsHandler = null) {
			
			this.DocumentPreprocessor = preprocessor;
			this.DocumentConverter = documentConverter;
			this.ConversionParameters = conversionParameters;
			this.EventsHandler = eventsHandler ?? new SilentEventsHandler();
			this.DocumentConverter.RemoveMessageListeners();
			this.DocumentConverter.AddProgressMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.EventsHandler.onProgressMessageReceived));
			this.DocumentConverter.AddFeedbackMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.EventsHandler.onFeedbackMessageReceived));
			this.DocumentConverter.AddFeedbackValidationListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.EventsHandler.onFeedbackValidationMessageReceived));
			this.DocumentConverter.DirectTransform = true;

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
		/// Preprocess a given document 
		/// </summary>
		/// <param name="inputPath"></param>
		/// <param name="resourceId"></param>
		/// <param name="withWordVisible">(defaults to true) Speicify if word must be visible or not during preprocessing</param>
		/// <returns>A document ready for conversion or null if an error has occured or if the preprocess has been canceled</returns>
		public DocumentParameters PreprocessDocument(string inputPath, string resourceId = null) {
			EventsHandler.onDocumentPreprocessingStart(inputPath);
			DocumentParameters result = new DocumentParameters(inputPath) {
				CopyPath = ConverterHelper.GetTempPath(inputPath, ".docx"),
				ResourceId = resourceId != null ? resourceId : null,
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
							EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Name validated, creating the working copy"));
							CurrentStatus  = DocumentPreprocessor.CreateWorkingCopy(ref preprocessedObject, ref result, EventsHandler);
							break;
						case ConversionStatus.CreatedWorkingCopy: // start processing shapes
							EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Working copy created, processing shapes"));
							CurrentStatus = DocumentPreprocessor.ProcessShapes(ref preprocessedObject, ref result, EventsHandler);
							break;
						case ConversionStatus.ProcessedShapes: // start processing math
							EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Shapes processed, processing mathml"));
							CurrentStatus = DocumentPreprocessor.ProcessEquations(ref preprocessedObject, ref result, EventsHandler);
							break;
						case ConversionStatus.ProcessedMathML: // finalize preprocessing
							EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("MathML processed"));
							CurrentStatus = DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
							break;
					}
				} while (CurrentStatus != ConversionStatus.Canceled &&
						CurrentStatus != ConversionStatus.Error &&
						CurrentStatus != ConversionStatus.PreprocessingSucceeded);
			} catch (Exception e) {
				EventsHandler.onPreprocessingError(inputPath, e.Message);
				throw;
			} finally {
				if (CurrentStatus != ConversionStatus.PreprocessingSucceeded) {
					DocumentPreprocessor.endPreprocessing(ref preprocessedObject, EventsHandler);
				}
			}
			if (CurrentStatus != ConversionStatus.PreprocessingSucceeded) {
				return null;
			}
			// Check for revisions
			if (result.HasRevisions) {
				result.TrackChanges = EventsHandler.AskForTrackConfirmation();
				// To be removed later, after moving TrackChanges eval from conversion to document param object
				ConversionParameters.TrackChanges = result.TrackChanges ? "Yes" : "No";
			} else {
				ConversionParameters.TrackChanges = "NoTrack";
			}

			// Only attempt to parse subdocument if no resourceId is provided or if nop attempt to parse was previously done 
			if (ConversionParameters.ParseSubDocuments == null || resourceId == null) {
				EventsHandler.onProgressMessageReceived(this, new DaisyEventArgs("Parsing subdocuments"));
				SubdocumentsList subDocList = SubdocumentsManager.FindSubdocuments(
					result.CopyPath,
					result.InputPath);
				result.HasSubDocuments = !subDocList.Empty;
				if (subDocList.Errors.Count > 0) {
					string errors = "Subdocuments convertion will be ignored due to the following errors found while extracting them:\r\n" + string.Join("\r\n", subDocList.Errors);
					EventsHandler.onPreprocessingError(result.InputPath, errors);
					ConversionParameters.ParseSubDocuments = false;
				} else if (result.HasSubDocuments) {
					ConversionParameters.ParseSubDocuments = false;
					if (EventsHandler != null) {
						ConversionParameters.ParseSubDocuments = EventsHandler.AskForTranslatingSubdocuments();
					}
					if (ConversionParameters.ParseSubDocuments) {
						foreach (SubdocumentInfo item in subDocList.Subdocuments) {
							if (CurrentStatus != ConversionStatus.Canceled) {
								DocumentParameters subDoc = null;
								try {
									subDoc = this.PreprocessDocument(item.FileName, item.RelationshipId);
								} catch (Exception e) {
									string errors = "Subdocuments convertion will be ignored due to the following errors found while preprocessing "+ item.FileName  +":\r\n" + e.Message;
									EventsHandler.onPreprocessingError(item.FileName, errors);
									ConversionParameters.ParseSubDocuments = false;
								}
								if (subDoc != null) {
									result.SubDocumentsToConvert.Add(subDoc);
								} else {
									// Cancel sub documents conversion
									result.SubDocumentsToConvert.Clear();
									break;
                                }
								
							} else {
								EventsHandler.onConversionCanceled();
								return null;
							}
                        }
					}
				} else {
					ConversionParameters.ParseSubDocuments = false;
				}
			}
			CurrentStatus = ConversionStatus.PreprocessingSucceeded;
			EventsHandler.onPreprocessingSuccess();
			return result;
			
		}



        /// <summary>
        /// Convert a single document to XML and apply post processing script on the result
        /// (Note that post-processing can be a conversion to another format)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        /// <param name="applyPostProcessing"></param>
        public ConversionResult Convert(DocumentParameters document, bool applyPostProcessing = true) {

			this.EventsHandler.onDocumentConversionStart(document, ConversionParameters);
			xmlConversionCancel = new CancellationTokenSource();
			CurrentStatus = ConversionStatus.HasStartedConversion;
			// If conversion is to be post processed output the xsl transfo to temp
			string outputDirectory = ConversionParameters.PostProcessor != null ?
				Path.GetTempPath() : (
					ConversionParameters.OutputPath.EndsWith(".xml") ?
						Directory.GetParent(ConversionParameters.OutputPath).FullName :
						ConversionParameters.OutputPath
				);

			// Rebuild and sanitize file name
			string outputFilename = (
				ConversionParameters.OutputPath.EndsWith(".xml") ?
					Path.GetFileName(ConversionParameters.OutputPath) :
					Path.GetFileNameWithoutExtension(document.InputPath) + ".xml"
				);

			string sanitizedName = outputFilename.Replace(",", "_");/*.Replace(" ", "_");*/

			// Rebuild output path
			document.OutputPath = Path.Combine(outputDirectory, sanitizedName);

			if (document.SubDocumentsToConvert.Count > 0 && ConversionParameters.ParseSubDocuments) {
                List<DocumentParameters> flattenList = new List<DocumentParameters> {
                    document
                };
                foreach (DocumentParameters subDocument in document.SubDocumentsToConvert) {
					flattenList.Add(subDocument);
				}
				this.Convert(flattenList, false);
			} else {
				try {
					conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
						return DocumentConverter.ConvertDocument(document, ConversionParameters);
					});
					conversionTask.Wait(xmlConversionCancel.Token);
					conversionTask.Dispose();
				} catch (OperationCanceledException) { 
				} catch (Exception e) {
					string message = e.Message;
					while(e.InnerException != null) {
						message += "\r\n - " + e.InnerException.Message;
						e = e.InnerException;
					}
					

					this.EventsHandler.OnUnknownError(message);
					return ConversionResult.Failed(message);
				}
				if (conversionTask.IsFaulted) {
					CurrentStatus = ConversionStatus.Error;
					Exception e = conversionTask.Exception;

					string message = e.Message;
					while (e.InnerException != null) {
						message += "\r\n - " + e.InnerException.Message;
						e = e.InnerException;
					}

					this.EventsHandler.OnUnknownError(message);
					return ConversionResult.Failed(message);
				}

			}
			if (CurrentStatus == ConversionStatus.Canceled) { // Conversion is aborted
				this.EventsHandler.onConversionCanceled();
				return ConversionResult.Cancel();
			}

			this.EventsHandler.onDocumentConversionSuccess(document, ConversionParameters);
			if (applyPostProcessing && ConversionParameters.PostProcessor != null) { // launch the pipeline post processing
				this.EventsHandler.onPostProcessingStart(ConversionParameters);
				try {
					ConversionParameters.PostProcessor.ExecuteScript(document.OutputPath);
				} catch (Exception e) {
					CurrentStatus = ConversionStatus.Error;
					this.EventsHandler.OnUnknownError(e.Message);
					return ConversionResult.Failed(e.Message);
				}
				this.EventsHandler.onPostProcessingSuccess(ConversionParameters);
			}
			return ConversionResult.Success();
		}

		/// <summary>
		/// Convert a list of document, merge them in a single XML file and apply post processing on the merged document
		/// </summary>
		/// <param name="documentLists">list of one or more document to convert</param>
		/// <param name="conversion">global conversion settings</param>
		/// <param name="applyPostProcessing">if true, post processing will be applied on the merge result</param>
		public ConversionResult Convert(List<DocumentParameters> documentLists, bool applyPostProcessing = true) {
            if (documentLists is null) {
                throw new ArgumentNullException(nameof(documentLists));
            }

            this.EventsHandler.onDocumentListConversionStart(documentLists, ConversionParameters);
			CurrentStatus = ConversionStatus.HasStartedConversion;
			string errors = "";

			try {
				XmlDocument mergeResult = new XmlDocument();
				if (documentLists.Count == 1) {
					return this.Convert(documentLists[0], false);
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
					foreach (DocumentParameters document in documentLists) {
						//document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
						// use memory temp for output
						document.OutputPath = Path.GetTempFileName();
						try {
							conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
								return DocumentConverter.ConvertDocument(document, ConversionParameters, mergeResult);
							});

							conversionTask.Wait(xmlConversionCancel.Token);
							if (conversionTask.IsCanceled) {
								this.EventsHandler.onConversionCanceled();
								return ConversionResult.Cancel();
							} else {
								mergeResult = conversionTask.Result;
							}
							conversionTask.Dispose();
						} catch (OperationCanceledException) {
							this.EventsHandler.onConversionCanceled();
							return ConversionResult.Cancel();
						} catch (AggregateException) {
							// Can be raise
						} catch (Exception e) {
							// TODO try to see if exception is raised by cancellation
							this.EventsHandler.OnUnknownError(document.InputPath + ": " + e.Message + "\r\n");
							return ConversionResult.Failed(document.InputPath + ": " + e.Message);
						}

						if (conversionTask.IsFaulted) {
							this.EventsHandler.OnUnknownError(conversionTask.Exception.Message);
							return ConversionResult.Failed(conversionTask.Exception.Message);
						}
						if (CurrentStatus == ConversionStatus.Canceled) {
							this.EventsHandler.onConversionCanceled();
							return ConversionResult.Cancel();
						}
						this.EventsHandler.onDocumentConversionSuccess(document, ConversionParameters);
					}
					DocumentConverter.finalizeAndSaveMergedDocument(mergeResult, ConversionParameters);
				}
			} catch (Exception e) {
				// Propagate unhandled exception
				throw new Exception(ResourceManager.GetString("TranslationFailed") + "\n"
					+ ResourceManager.GetString("WellDaisyFormat") + "\n" + " \""
					+ errors + "\n" + "Crictical issue:" + "\n" + e.Message + "\n", e);

			}
			if (DocumentConverter.ValidationErrors.Count > 0) {
				this.EventsHandler.OnValidationErrors(DocumentConverter.ValidationErrors, ConversionParameters.OutputPath );
				return ConversionResult.FailedOnValidation(
					string.Join(
						"\r\n",
						DocumentConverter.ValidationErrors.Select(
							error => error.ToString()
						).ToArray()
					)
				);
			} else if (CurrentStatus == ConversionStatus.Canceled) {
				return ConversionResult.Cancel();
			} else {
				this.EventsHandler.onDocumentListConversionSuccess(documentLists, ConversionParameters);
				
				if (DocumentConverter.LostElements.Count > 0) {
					ArrayList unconvertedElements = new ArrayList();
                    foreach (KeyValuePair<string, List<string> > lostElementForFile in DocumentConverter.LostElements) {
						if(lostElementForFile.Value.Count > 0) {
							string lostElements = lostElementForFile.Key + ":\r\n";
                            foreach (string lostElement in lostElementForFile.Value) {
								lostElements += " - " + lostElement + "\r\n";
							}
							unconvertedElements.Add(lostElements);
                        }
                    }
					this.EventsHandler.OnLostElements(ConversionParameters.OutputPath, unconvertedElements);
					applyPostProcessing = this.EventsHandler.IsContinueDTBookGenerationOnLostElements();
				}
				if (applyPostProcessing && ConversionParameters.PostProcessor != null) {
					// If post processing is requested (and can be applied even if lost elements are found)
					// Launch the post processing pipeline sript 
					// (cleaning or converting DTBook to another format Like a DAISY book)
					this.EventsHandler.onPostProcessingStart(ConversionParameters);
					try {
						ConversionParameters.PostProcessor.ExecuteScript(ConversionParameters.OutputPath);
					} catch (Exception e) {
						this.EventsHandler.OnUnknownError(e.Message);
						return ConversionResult.Failed(e.Message);
					}
					this.EventsHandler.onPostProcessingSuccess(ConversionParameters);
				}
			}
			return ConversionResult.Success();

		}


		/// <summary>
		/// 
		/// </summary>
		protected void RequestConversionCancel() {
			if (xmlConversionCancel != null) {
				xmlConversionCancel.Cancel();
			}
			CurrentStatus = ConversionStatus.Canceled;
			
		}


	}

}
