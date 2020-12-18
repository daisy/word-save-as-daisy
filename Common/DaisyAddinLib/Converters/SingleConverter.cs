using System;
using System.Collections;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Windows.Forms;

namespace Daisy.DaisyConverter.DaisyConverterLib.Converters
{
	/// <summary>
	/// Implementation of convertation.
	/// </summary>
	public class SingleConverter
	{
		public ScriptParser ScriptToExecute { get; set; }
		private AbstractConverter converter;
		ChainResourceManager resourceManager;
		string validationErrorMsg = "";
		private bool continueDTBookGeneration = true;
		int flag;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="converter">An implementation of AbstractConverter</param>
		public SingleConverter(AbstractConverter converter)
			: this(converter, null)
		{

		}

		public SingleConverter(AbstractConverter converter, ScriptParser scriptToExecute)
		{
			ScriptToExecute = scriptToExecute;
			this.converter = converter;
			this.resourceManager = new ChainResourceManager();
			// Add a default resource managers (for common labels)
			this.resourceManager.Add(new System.Resources.ResourceManager("DaisyAddinLib.resources.Labels",
				Assembly.GetExecutingAssembly()));
		}

		/// <summary>
		/// Override default resource manager.
		/// </summary>
		public System.Resources.ResourceManager OverrideResourceManager
		{
			set { this.resourceManager.Add(value); }
		}


		/// <summary>
		/// Retrieve the label associated to the specified key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public string GetString(string key)
		{
			return this.resourceManager.GetString(key);
		}

		/// <summary>
		/// Resource manager
		/// </summary>
		public System.Resources.ResourceManager ResManager
		{
			get { return this.resourceManager; }
		}

		/// <summary>
		/// Function to translate all Sub documents 
		/// </summary>
		/// <param name="tempInputFile">Duplicate for the Input File</param>
		/// <param name="temp_InputFile">Duplicate for the Input File</param>
		/// <param name="inputFile">Input file to be Translated</param>
		/// <param name="outputfilepath">Path of the Output file</param>
		/// <param name="HTable">Document properties</param>
		public void OoxToDaisyOwn(String tempInputFile, String temp_InputFile, String inputFile, String outputfilepath, Hashtable HTable, string control, Hashtable listMathMl, string output_Pipeline)
		{
			try
			{
				SubdocumentsList subdocuments = SubdocumentsManager.FindSubdocuments(temp_InputFile, inputFile);
				ArrayList notTranslatedDoc = subdocuments.GetNotTraslatedSubdocumentsNames();

				ArrayList subList = new ArrayList();
				subList.Add(tempInputFile + "|Master");
				subList.AddRange(subdocuments.GetSubdocumentsNameWithRelationship());
				int subCount = subdocuments.SubdocumentsCount + 1;

				//Checking whether any original or Subdocumets is already Open or not
				string resultOpenSub = CheckFileOPen(subList);
				if (resultOpenSub != "notopen")
				{
					OnUnknownError(resourceManager.GetString("OpenSubOwn"));
					return;
				}

				//Checking whether Sub documents are Simple documents or a Master document
				string resultSub = SubdocumentsManager.CheckingSubDocs(subdocuments.GetSubdocumentsNameWithRelationship());
				if (resultSub != "simple")
				{
					OnUnknownError(resourceManager.GetString("AddSimpleMasterSub"));
					return;
				}

				if (subCount != subList.Count)
				{
					OnUnknownError(this.resourceManager.GetString("ProblemMasterSub"));
					return;
				}

				OoxToDaisySub(outputfilepath, subList, HTable, tempInputFile, control, listMathMl, output_Pipeline,
							  notTranslatedDoc);

			}
			catch (Exception e)
			{
				AddinLogger.Error(e);
				OnUnknownError(resourceManager.GetString("TranslationFailed") + "\n" + this.resourceManager.GetString("WellDaisyFormat") + "\n" + " \"" + Path.GetFileName(tempInputFile) + "\"\n" + validationErrorMsg + "\n" + "Problem is:" + "\n" + e.Message + "\n");
			}
		}

		/// <summary>
		/// Function to translate all Bunch of documents  selected by the user
		/// </summary>
		/// <param name="outputfilepath">Path of the output File</param>
		/// <param name="subList">List of All documents</param>
		/// <param name="category">whether master/sub doc or Bunch of Docs</param>
		/// <param name="table">Document Properties</param>
		public bool OoxToDaisySub(string outputfilepath, ArrayList subList, string category, Hashtable table, string control, Hashtable MultipleMathMl, string output_Pipeline)
		{
			flag = 0;
			validationErrorMsg = "";
			using (Progress progress = new Progress(this.converter, this.resourceManager, outputfilepath, subList, category, table, control, MultipleMathMl, output_Pipeline))
			{
				DialogResult progressDialogResult = progress.ShowDialog();
				if (progressDialogResult == DialogResult.OK)
				{
					if (!string.IsNullOrEmpty(progress.ValidationError))
					{
						validationErrorMsg = progress.ValidationError;
						OnMasterSubValidationError(progress.ValidationError);
					}
					else if (progress.HasLostElements)
					{
						OnLostElements(string.Empty, outputfilepath, progress.LostElements);
						flag = 1;
						if (!(AddInHelper.IsSingleDaisyFromMultipleButton(control) && ScriptToExecute == null))
							continueDTBookGeneration = IsContinueDTBookGenerationOnLostElements();
					}
					else
					{
						if (AddInHelper.IsSingleDaisyFromMultipleButton(control) && ScriptToExecute == null)
						{
							OnSuccess();
						}
						flag = 1;
					}
				}
				else if (progressDialogResult == DialogResult.Cancel)
				{
					DeleteDTBookFilesIfExists(outputfilepath);
				}
				else if (!string.IsNullOrEmpty(progress.ValidationError))
				{
					OnMasterSubValidationError(progress.ValidationError);
				}
			}

			DeleteTemporaryImages();

			return (flag == 1) && continueDTBookGeneration;
		}


		/// <summary>
		/// Function to translate all sub documents in the Master Document
		/// </summary>
		/// <param name="outputfilepath">Path of the output File</param>
		/// <param name="subList">List of All documents</param>
		/// <param name="HTable">Document Properties</param>
		/// <param name="tempInputFile">Duplicate for the Input File</param>
		public void OoxToDaisySub(String outputfilepath, ArrayList subList, Hashtable HTable, String tempInputFile, string control, Hashtable listMathMl, string output_Pipeline, ArrayList notTranslatedDoc)
		{
			flag = 0;
			validationErrorMsg = "";
			using (Progress progress = new Progress(this.converter, this.resourceManager, outputfilepath, subList, HTable, tempInputFile, control, listMathMl, output_Pipeline))
			{
				DialogResult dr = progress.ShowDialog();
				if (dr == DialogResult.OK)
				{
					validationErrorMsg = progress.ValidationError;
					String messageDocsSkip = DocumentSkipped(notTranslatedDoc);
					if (!string.IsNullOrEmpty(validationErrorMsg))
					{
						validationErrorMsg = validationErrorMsg + messageDocsSkip;
						OnMasterSubValidationError(validationErrorMsg);
					}
					else if (progress.HasLostElements)
					{
						OnLostElements(string.Empty, outputfilepath + "\\1.xml", progress.LostElements);

						if (AddInHelper.PipelineIsInstalled() && 
								AddInHelper.buttonIsSingleWordToXMLConversion(control) && 
								ScriptToExecute != null && 
								IsContinueDTBookGenerationOnLostElements())
						{
							try
							{
								ExecuteScript(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
							}
							catch (Exception e)
							{
								AddinLogger.Error(e);
								OnUnknownError(e.Message);
							}
						}
					}
					else
					{
						if (!string.IsNullOrEmpty(messageDocsSkip))
						{
							OnSuccessMasterSubValidation(ResManager.GetString("SucessLabel") + messageDocsSkip);
						}
						else
						{
							if (AddInHelper.IsSingleDaisyTranslate(control) && ScriptToExecute == null)
							{
								OnSuccess();
							}
							else ExecuteScript(outputfilepath + "\\" + Path.GetFileNameWithoutExtension(tempInputFile) + ".xml");
						}
					}
				}
				else if (dr == DialogResult.Cancel)
				{
					DeleteDTBookFilesIfExists(outputfilepath);
				}
				else
				{
					validationErrorMsg = progress.ValidationError;
					if (!string.IsNullOrEmpty(validationErrorMsg))
					{
						OnMasterSubValidationError(validationErrorMsg);
					}
				}
			}
			DeleteTemporaryImages();
		}

		private ConvertResult ConvertOoxToDaisy(string inputFile, string outputFile, Hashtable listMathMl, Hashtable table, string control, string output_Pipeline)
		{
			try
			{
				using (ConverterForm form = new ConverterForm(this.converter, inputFile, outputFile, this.resourceManager, true, listMathMl, table, control, output_Pipeline))
				{
					if (DialogResult.OK != form.ShowDialog())
						return ConvertResult.Cancel();

					if (!String.IsNullOrEmpty(form.ValidationError))
					{
						OnValidationError(form.ValidationError, inputFile, outputFile);
						return ConvertResult.ValidationError(form.ValidationError);
					}

					if (form.HasLostElements) {
						OnLostElements(inputFile, outputFile, form.LostElements);

						if (!(AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) && IsContinueDTBookGenerationOnLostElements()) {
							ExecuteScript(outputFile);
						}
					} else if (AddInHelper.IsSingleDaisyTranslate(control) && this.ScriptToExecute == null) {
						OnSuccess();
					} else {
						ExecuteScript(outputFile);
					}
				}
			}
			catch (IOException e)
			{
				// this is meant to catch "file already accessed by another process", though there's no .NET fine-grain exception for this.
				AddinLogger.Error(e);
				OnUnknownError("UnableToCreateOutputLabel", e.Message);
				return ConvertResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			}
			catch (Exception e)
			{
				AddinLogger.Error(e);
				OnUnknownError("DaisyUnexpectedError", e.GetType() + ": " + e.Message + " (" + e.StackTrace + ")");

				if (File.Exists(outputFile))
				{
					File.Delete(outputFile);
				}

				return ConvertResult.UnknownError(e.Message + Environment.NewLine + e.StackTrace);
			}

			return ConvertResult.Success();
		}

		public ConvertResult OoxToDaisy(string inputFile, string outputFile, Hashtable listMathMl, Hashtable table, string control, string output_Pipeline)
		{
			ConvertResult result = ConvertOoxToDaisy(inputFile, outputFile, listMathMl, table, control, output_Pipeline);

			DeleteTemporaryImages();

			return result;
		}

		#region help methods

		private void ExecuteScript(string inputDaisyXmlPath)
		{
			if (ScriptToExecute != null)
			{
				try
				{
					ScriptToExecute.ExecuteScript(inputDaisyXmlPath);
				}
				catch (Exception e)
				{
					AddinLogger.Error(e);
					OnUnknownError(e.Message);
				}
			}
		}

		private void DeleteDTBookFilesIfExists(string outputfilepath)
		{
			if (File.Exists(outputfilepath + "\\dtbookbasic.css"))
				File.Delete(outputfilepath + "\\dtbookbasic.css");
			if (File.Exists(outputfilepath + "\\dtbook-2005-3.dtd"))
				File.Delete(outputfilepath + "\\dtbook-2005-3.dtd");
		}

		private string DocumentSkipped(ArrayList notTranslatedDoc)
		{
			string message = "";
			if (notTranslatedDoc.Count != 0)
			{
				for (int i = 0; i < notTranslatedDoc.Count; i++)
					message = message + Convert.ToString(i + 1) + ". " + notTranslatedDoc[i].ToString() + "\n";

				message = "\n\n" + "Files which are not in Word 2007 format are skipped during Translation:" + "\n" + message;
			}
			return message;
		}

		private void DeleteTemporaryImages()
		{
			string[] files = Directory.GetFiles(AddInHelper.AppDataSaveAsDAISYDirectory);
			foreach (string file in files)
			{
				if (file.Contains(".jpg") || file.Contains(".JPG") || file.Contains(".PNG") || file.Contains(".png"))
				{
					File.Delete(file);
				}
			}
		}

		#endregion

		#region Multiple OOXML

		/// <summary>
		/// Function to check Whether Docs are already open or not
		/// </summary>
		/// <param name="listSubDocs">List of Sub Documents</param>
		/// <returns>Message whether any doc in List are open or not</returns>
		public static string CheckFileOPen(ArrayList listSubDocs)
		{
			String resultSubDoc = "notopen";
			for (int i = 0; i < listSubDocs.Count; i++)
			{
				string[] splt = listSubDocs[i].ToString().Split('|');
				try
				{
					Package pack;
					pack = Package.Open(splt[0].ToString(), FileMode.Open, FileAccess.ReadWrite);
					pack.Close();
				}
				catch (Exception e)
				{
					AddinLogger.Error(e);
					resultSubDoc = "open";
				}
			}
			return resultSubDoc;
		}

		#endregion

		#region virtual methods

		protected virtual void OnSuccessMasterSubValidation(string message)
		{

		}

		protected virtual void OnUnknownError(string error)
		{

		}

		protected virtual void OnUnknownError(string title, string details)
		{
		}

		protected virtual void OnValidationError(string error, string inputFile, string outputFile)
		{
		}

		protected virtual void OnLostElements(string inputFile, string outputFile, ArrayList elements)
		{
		}

		protected virtual bool IsContinueDTBookGenerationOnLostElements()
		{
			return true;
		}

		protected virtual void OnSuccess()
		{
		}

		protected virtual void OnMasterSubValidationError(string error)
		{

		}

		#endregion
	}
}