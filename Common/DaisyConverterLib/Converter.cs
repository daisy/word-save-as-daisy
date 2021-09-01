using Daisy.SaveAsDAISY.Conversion.Events;
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
		protected WordToDTBookXMLTransform documentConverter;
        protected ConversionParameters conversion;



        protected ChainResourceManager resourceManager;

		protected string validationErrorMsg = "";

		protected int flag;

		private Task<XmlDocument> conversionTask;
		private CancellationTokenSource xmlConversionCancel;



		public ConversionStatus CurrentStatus = ConversionStatus.None;

        /// <summary>
        /// Override default resource manager.
        /// </summary>
        public System.Resources.ResourceManager OverrideResourceManager {
			set { this.resourceManager.Add(value); }
		}


		/// <summary>
		/// Retrieve the label associated to the specified key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public string GetString(string key) {
			return this.resourceManager.GetString(key);
		}

		/// <summary>
		/// Resource manager
		/// </summary>
		public System.Resources.ResourceManager ResManager {
			get { return this.resourceManager; }
		}

        //public ConversionParameters Conversion { get => conversion; set => conversion = value; }

        /// <summary>
        /// Events handler class
        /// </summary>
        protected IConversionEventsHandler eventsHandler;

		/// <summary>
		/// Document preprocessor, that depends on word interop version
		/// </summary>
		protected IDocumentPreprocessor documentPreprocessor;

		public Converter(IDocumentPreprocessor preprocessor, WordToDTBookXMLTransform documentConverter, ConversionParameters conversionParameters, IConversionEventsHandler eventsHandler = null) {
			
			this.documentPreprocessor = preprocessor;
			this.documentConverter = documentConverter;
			this.conversion = conversionParameters;
			this.eventsHandler = eventsHandler ?? new SilentEventsHandler();
			this.documentConverter.RemoveMessageListeners();
			this.documentConverter.AddProgressMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onProgressMessageReceived));
			this.documentConverter.AddFeedbackMessageListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onFeedbackMessageReceived));
			this.documentConverter.AddFeedbackValidationListener(new WordToDTBookXMLTransform.XSLTMessagesListener(this.eventsHandler.onFeedbackValidationMessageReceived));
			this.documentConverter.DirectTransform = true;

			this.resourceManager = new ChainResourceManager();
			// Add a default resource managers (for common labels)
			this.resourceManager.Add(
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
		/// <returns>A document ready for conversion or null if an error has occured or if the preprocess has been canceled</returns>
		public DocumentParameters preprocessDocument(string inputPath, string resourceId = null) {
			eventsHandler.onDocumentPreprocessingStart(inputPath);
			DocumentParameters result = new DocumentParameters(inputPath) {
				CopyPath = ConverterHelper.GetTempPath(inputPath, ".docx"),
				ResourceId = resourceId != null ? resourceId : null
			};
			// dot not make visible subdocuments (documents with resource Id assigned)
			object preprocessedObject = documentPreprocessor.startPreprocessing(result, eventsHandler);
			try {
				do {
					switch (CurrentStatus) {
						case ConversionStatus.None: // Starting by validating file name
							CurrentStatus = documentPreprocessor.ValidateName(ref preprocessedObject, conversion.NameValidator, eventsHandler);
							break;
						case ConversionStatus.ValidatedName: // make the working copy
							CurrentStatus  = documentPreprocessor.CreateWorkingCopy(ref preprocessedObject, ref result, eventsHandler);
							break;
						case ConversionStatus.CreatedWorkingCopy: // start processing shapes
							CurrentStatus = documentPreprocessor.ProcessShapes(ref preprocessedObject, ref result, eventsHandler);
							break;
						case ConversionStatus.ProcessedShapes: // start processing math
							CurrentStatus = documentPreprocessor.ProcessEquations(ref preprocessedObject, ref result, eventsHandler);
							break;
						case ConversionStatus.ProcessedMathML: // finalize preprocessing
							CurrentStatus = documentPreprocessor.endPreprocessing(ref preprocessedObject, eventsHandler);
							break;
					}
				} while (CurrentStatus != ConversionStatus.Canceled &&
						CurrentStatus != ConversionStatus.Error &&
						CurrentStatus != ConversionStatus.PreprocessingSucceeded);
			} catch (Exception e) {
				eventsHandler.onPreprocessingError(inputPath, e.Message);
				throw e;
			} finally {
				if (CurrentStatus != ConversionStatus.PreprocessingSucceeded) {
					documentPreprocessor.endPreprocessing(ref preprocessedObject, eventsHandler);
				}
			}
			if (CurrentStatus != ConversionStatus.PreprocessingSucceeded) {
				return null;
			}
			
			// Check for revisions
			if (result.HasRevisions) {
				result.TrackChanges = eventsHandler.AskForTrackConfirmation();
				// To be removed later, after moving TrackChanges eval from conversion to document param object
				conversion.TrackChanges = result.TrackChanges ? "Yes" : "No";
			} else {
				conversion.TrackChanges = "NoTrack";
			}

			// Only attempt to parse subdocument if no resourceId is provided or if nop attempt to parse was previously done 
			if (conversion.ParseSubDocuments == null || resourceId == null) {
				SubdocumentsList subDocList = SubdocumentsManager.FindSubdocuments(
					result.CopyPath,
					result.InputPath);
				result.HasSubDocuments = !subDocList.Empty;
				if (subDocList.Errors.Count > 0) {
					string errors = "Subdocuments convertion will be ignored due to the following errors found while extracting them:\r\n" + string.Join("\r\n", subDocList.Errors);
					eventsHandler.onPreprocessingError(result.InputPath, errors);
					conversion.ParseSubDocuments = "No";
				} else if (result.HasSubDocuments) {
					conversion.ParseSubDocuments = "No";
					if (eventsHandler != null) {
						conversion.ParseSubDocuments = eventsHandler.AskForTranslatingSubdocuments() ? "Yes" : "No";
					}
					if (conversion.ParseSubDocuments == "Yes") {
						foreach (SubdocumentInfo item in subDocList.Subdocuments) {
							if (CurrentStatus != ConversionStatus.Canceled) {
								DocumentParameters subDoc = null;
								try {
									subDoc = this.preprocessDocument(item.FileName, item.RelationshipId);
								} catch (Exception e) {
									string errors = "Subdocuments convertion will be ignored due to the following errors found while preprocessing "+ item.FileName  +":\r\n" + e.Message;
									eventsHandler.onPreprocessingError(item.FileName, errors);
									conversion.ParseSubDocuments = "No";
								}
								if (subDoc != null) {
									result.SubDocumentsToConvert.Add(subDoc);
								} else {
									// Cancel sub documents conversion
									result.SubDocumentsToConvert.Clear();
									break;
                                }
								
							} else {
								eventsHandler.onConversionCanceled();
								return null;
							}
                        }
					}
				} else {
					conversion.ParseSubDocuments = "NoMasterSub";
				}
			}
			CurrentStatus = ConversionStatus.PreprocessingSucceeded;
			eventsHandler.onPreprocessingSuccess();
			return result;
			
		}



        /// <summary>
        /// Convert a single document to XML and apply post processing script on the result
        /// (Note that post-processing can be a conversion to another format)
        /// </summary>
        /// <param name="document"></param>
        /// <param name="conversion"></param>
        /// <param name="applyPostProcessing"></param>
        public ConversionResult convert(DocumentParameters document, bool applyPostProcessing = true) {

			this.eventsHandler.onDocumentConversionStart(document, conversion);
			xmlConversionCancel = new CancellationTokenSource();
			CurrentStatus = ConversionStatus.HasStartedConversion;
			// If conversion is to be post processed output the xsl transfo to temp
			string outputDirectory = conversion.ScriptPath != null ?
				Path.GetTempPath() : (
					conversion.OutputPath.EndsWith(".xml") ?
						Directory.GetParent(conversion.OutputPath).FullName :
						conversion.OutputPath
				);

			// Rebuild and sanitize file name
			string outputFilename = (
				conversion.OutputPath.EndsWith(".xml") ?
					Path.GetFileName(conversion.OutputPath) :
					Path.GetFileNameWithoutExtension(document.InputPath) + ".xml"
				);

			string sanitizedName = outputFilename.Replace(",", "_");/*.Replace(" ", "_");*/

			// Rebuild output path
			document.OutputPath = Path.Combine(outputDirectory, sanitizedName);

			if (document.SubDocumentsToConvert.Count > 0 && conversion.ParseSubDocuments.ToLower() == "yes") {
                List<DocumentParameters> flattenList = new List<DocumentParameters> {
                    document
                };
                foreach (DocumentParameters subDocument in document.SubDocumentsToConvert) {
					flattenList.Add(subDocument);
				}
				this.convert(flattenList, false);
			} else {
				try {
					conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
						return documentConverter.ConvertDocument(document, conversion);
					});
					conversionTask.Wait(xmlConversionCancel.Token);
					conversionTask.Dispose();
				} catch (OperationCanceledException) { 
				} catch (Exception e) {
					this.eventsHandler.OnUnknownError(e.Message);
					return ConversionResult.Failed(e.Message);
				}
				if (conversionTask.IsFaulted) {
					CurrentStatus = ConversionStatus.Error;
					this.eventsHandler.OnUnknownError(conversionTask.Exception.Message);
					return ConversionResult.Failed(conversionTask.Exception.Message);
				}

			}
			if (CurrentStatus == ConversionStatus.Canceled) { // Conversion is aborted
				this.eventsHandler.onConversionCanceled();
				return ConversionResult.Cancel();
			}

			this.eventsHandler.onDocumentConversionSuccess(document, conversion);
			if (applyPostProcessing && conversion.PostProcessSettings != null) { // launch the pipeline post processing
				this.eventsHandler.onPostProcessingStart(conversion);
				try {
					conversion.PostProcessSettings.ExecuteScript(document.OutputPath);
				} catch (Exception e) {
					CurrentStatus = ConversionStatus.Error;
					this.eventsHandler.OnUnknownError(e.Message);
					return ConversionResult.Failed(e.Message);
				}
				this.eventsHandler.onPostProcessingSuccess(conversion);
			}
			return ConversionResult.Success();
		}

		/// <summary>
		/// Convert a list of document, merge them in a single XML file and apply post processing on the merged document
		/// </summary>
		/// <param name="documentLists">list of one or more document to convert</param>
		/// <param name="conversion">global conversion settings</param>
		/// <param name="applyPostProcessing">if true, post processing will be applied on the merge result</param>
		public ConversionResult convert(List<DocumentParameters> documentLists, bool applyPostProcessing = true) {
			this.eventsHandler.onDocumentListConversionStart(documentLists, conversion);
			CurrentStatus = ConversionStatus.HasStartedConversion;
			string errors = "";

			try {
				XmlDocument mergeResult = new XmlDocument();
				if (documentLists.Count == 1) {
					return this.convert(documentLists[0], false);
				} else {
					string outputDirectory = conversion.OutputPath.EndsWith(".xml") ?
						Directory.GetParent(conversion.OutputPath).FullName :
						conversion.OutputPath;
					// Rebuild and sanitize file name
					string outputFilename = (
						conversion.OutputPath.EndsWith(".xml") ?
							Path.GetFileName(conversion.OutputPath) :
							Path.GetFileNameWithoutExtension(documentLists[0].InputPath) + ".xml"
						).Replace(" ", "_").Replace(",", "_");

					// Rebuild output path based on first document 
					conversion.OutputPath = Path.Combine(outputDirectory, outputFilename);
					foreach (DocumentParameters document in documentLists) {
						//document.OutputPath = outputDirectory + "\\" + Path.GetFileNameWithoutExtension(document.InputPath) + ".xml";
						// use memory temp for output
						document.OutputPath = Path.GetTempFileName();
						try {
							conversionTask = Task<XmlDocument>.Factory.StartNew(() => {
								return documentConverter.ConvertDocument(document, conversion, mergeResult);
							});

							conversionTask.Wait(xmlConversionCancel.Token);
							if (conversionTask.IsCanceled) {
								this.eventsHandler.onConversionCanceled();
								return ConversionResult.Cancel();
							} else {
								mergeResult = conversionTask.Result;
							}
							conversionTask.Dispose();
						} catch (OperationCanceledException) {
							this.eventsHandler.onConversionCanceled();
							return ConversionResult.Cancel();
						} catch (AggregateException) {
							// Can be raise
						} catch (Exception e) {
							// TODO try to see if exception is raised by cancellation
							this.eventsHandler.OnUnknownError(document.InputPath + ": " + e.Message + "\r\n");
							return ConversionResult.Failed(document.InputPath + ": " + e.Message);
						}

						if (conversionTask.IsFaulted) {
							this.eventsHandler.OnUnknownError(conversionTask.Exception.Message);
							return ConversionResult.Failed(conversionTask.Exception.Message);
						}
						if (CurrentStatus == ConversionStatus.Canceled) {
							this.eventsHandler.onConversionCanceled();
							return ConversionResult.Cancel();
						}
					}
					documentConverter.finalizeAndSaveMergedDocument(mergeResult, conversion);
				}
			} catch (Exception e) {
				// Propagate unhandled exception
				throw new Exception(resourceManager.GetString("TranslationFailed") + "\n"
					+ resourceManager.GetString("WellDaisyFormat") + "\n" + " \""
					+ errors + "\n" + "Crictical issue:" + "\n" + e.Message + "\n");

			}
			if (documentConverter.ValidationErrors.Count > 0) {
				this.eventsHandler.OnValidationErrors(documentConverter.ValidationErrors, conversion.OutputPath );
				return ConversionResult.FailedOnValidation(
					string.Join(
						"\r\n",
						documentConverter.ValidationErrors.Select(
							error => error.ToString()
						).ToArray()
					)
				);
			} else if (CurrentStatus == ConversionStatus.Canceled) {
				return ConversionResult.Cancel();
			} else {
				this.eventsHandler.onDocumentListConversionSuccess(documentLists, conversion);
				
				if (documentConverter.LostElements.Count > 0) {
					ArrayList unconvertedElements = new ArrayList();
                    foreach (KeyValuePair<string, List<string> > lostElementForFile in documentConverter.LostElements) {
						if(lostElementForFile.Value.Count > 0) {
							string lostElements = lostElementForFile.Key + ":\r\n";
                            foreach (string lostElement in lostElementForFile.Value) {
								lostElements += " - " + lostElement + "\r\n";
							}
							unconvertedElements.Add(lostElements);
                        }
                    }
					this.eventsHandler.OnLostElements(conversion.OutputPath, unconvertedElements);
					applyPostProcessing = this.eventsHandler.IsContinueDTBookGenerationOnLostElements();
				}
				if (applyPostProcessing && conversion.PostProcessSettings != null) {
					// If post processing is requested (and can be applied even if lost elements are found)
					// Launch the post processing pipeline sript 
					// (cleaning or converting DTBook to another format Like a DAISY book)
					this.eventsHandler.onPostProcessingStart(conversion);
					try {
						conversion.PostProcessSettings.ExecuteScript(conversion.OutputPath);
					} catch (Exception e) {
						this.eventsHandler.OnUnknownError(e.Message);
						return ConversionResult.Failed(e.Message);
					}
					this.eventsHandler.onPostProcessingSuccess(conversion);
				}
			}
			return ConversionResult.Success();

		}


		/// <summary>
		/// 
		/// </summary>
		protected void requestConversionCancel() {
			if (xmlConversionCancel != null) {
				xmlConversionCancel.Cancel();
			}
			CurrentStatus = ConversionStatus.Canceled;
			
		}


	}

}
