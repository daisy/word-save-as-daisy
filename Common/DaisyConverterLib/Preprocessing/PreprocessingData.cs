using Daisy.SaveAsDAISY.Conversion.Events;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Xml;

namespace Daisy.SaveAsDAISY.Conversion {

	// TODO remove this class, everything is now stored and habdled by DocumentParameters, ConversionParameters and the converter class
	public class PreprocessingData
	{

		IConversionEventsHandler eventsHandler;


		public PreprocessingData(IConversionEventsHandler eventsHandler = null)
		{
			this.eventsHandler = eventsHandler;
		}

		public PreprocessingData(IConversionEventsHandler eventsHandler, string wordVersion,  Pipeline pipeline = null, string pipelineScriptKey = "" ) : base() {
			this.eventsHandler = eventsHandler;
			conversionParameters.Version = wordVersion;
			FileInfo postprocessScriptFile;
			if (pipeline == null) {
				conversionParameters.ScriptPath = null;
			} else if (pipelineScriptKey != "") {
				conversionParameters.ScriptPath = pipeline.ScriptsInfo[pipelineScriptKey].FullName;
			} else if (pipeline.ScriptsInfo.TryGetValue("_postprocess", out postprocessScriptFile)) {
				// Note : adding a default postprocess script for dtbook pipeline special treatment
				// This script is alledgedly not visible to users
				conversionParameters.ScriptPath = postprocessScriptFile.FullName;
			} else conversionParameters.ScriptPath = null;
		}

		public static PreprocessingData Failed(string error)
		{
			return new PreprocessingData() { IsSuccess = false, IsCanceled = false, LastMessage = error };
		}

		public static PreprocessingData Canceled(string message) {
			return new PreprocessingData() { IsSuccess = false, IsCanceled = true, LastMessage = message };
		}

		// converter settings, to be initialized when needed
		private ConversionParameters _conversionParameters = null;
		public ConversionParameters conversionParameters { 
			get {
				if (_conversionParameters == null) {
					//_conversionParameters = new ConversionParameters(eventsHandler ?? new SilentEventsHandler());
					
				}
				return _conversionParameters;
			} 
		}
		/// <summary>
		/// List of document that have been preprocessed
		/// </summary>
		private List<DocumentParameters> documents = null;
		public List<DocumentParameters> Documents {
			get {
				if (documents == null) {
					documents = new List<DocumentParameters>();
				}
				return documents;
			}
		}

		public bool IsSuccess { get; set; }

		public bool IsCanceled { get; set; }
		public string LastMessage { get; set; }

		/// <summary>
		/// Function to generate Random ID
		/// </summary>
		/// <returns></returns>
		static public long GenerateId() {
			byte[] buffer = Guid.NewGuid().ToByteArray();
			return BitConverter.ToInt64(buffer, 0);
		}

		public static string requestSubDocumentsConversion(string tempInput, IConversionEventsHandler eventsHandler) {

			const string wordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
			PackageRelationship packRelationship = null;

			Package pack = Package.Open(tempInput, FileMode.Open, FileAccess.ReadWrite);

			foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(wordRelationshipType)) {
				packRelationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(packRelationship.SourceUri, packRelationship.TargetUri);
			PackagePart mainPartxml = pack.GetPart(partUri);
			int cnt = 0;
			foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships()) {
				packRelationship = searchRelation;
				if (packRelationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument") {
					if (packRelationship.TargetMode.ToString() == "External") {
						cnt++;
					}
				}
			}
			pack.Close();
			return cnt == 0 ? "NoMasterSub" : (eventsHandler.AskForTranslatingSubdocuments() ? "Yes" : "No");
		}
	}
}