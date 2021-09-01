using System;
using System.Collections;
using System.IO;
using System.IO.Packaging;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// Provides methos for working with MS Word subdocuments.
	/// </summary>
	public class SubdocumentsManager
	{
		const string WordRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
		const string WordSubdocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument";
		private const string LocalSettingsTemp = "/Local Settings/Temp/";

		

		/// <summary>
		/// Find all subdocuments.
		/// </summary>
		/// <param name="inputPath">Path to the word document.</param>
		/// <returns>Collection of the subdocuments.</returns>
		public static SubdocumentsList FindSubdocuments(string inputPath)
		{
			return FindSubdocuments(inputPath, inputPath);
		}

		/// <summary>
		/// Find all subdocuments.
		/// The Errors array is filled if errors are found
		/// </summary>
		/// <param name="tempInputPath">Path to the copy of the word document.</param>
		/// <param name="originalInputPath">Path to the original word document.</param>
		/// <returns>Collection of the subdocuments.</returns>
		public static SubdocumentsList FindSubdocuments(string tempInputPath, string originalInputPath)
		{
			PackageRelationship relationship = null;
			SubdocumentsList result = new SubdocumentsList();

			Package packDoc = Package.Open(tempInputPath, FileMode.Open, FileAccess.ReadWrite);

			foreach (PackageRelationship searchRelation in packDoc.GetRelationshipsByType(WordRelationshipType))
			{
				relationship = searchRelation;
				break;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			PackagePart mainPartxml = packDoc.GetPart(partUri);

			foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
			{
				relationship = searchRelation;
				if (IsExternalSubdocumentRelationship(relationship)) {
					bool subDocumentFound = IsMsWordSubdocumentRelationship(relationship);
					result.SubdocumentsCount += subDocumentFound ? 1 : 0;
					String filePath = GetRealFilePath(relationship.TargetUri.ToString(), originalInputPath);
					if (!string.IsNullOrEmpty(filePath)) {
						if (subDocumentFound) {
							result.Subdocuments.Add(new SubdocumentInfo(filePath, relationship.Id));
						}
						else {
							result.NotTranslatedSubdocuments.Add(new SubdocumentInfo(filePath, relationship.Id));
						}
					} else {
						result.Errors.Add(
							"The subdocument " + originalInputPath + "\\" + relationship.TargetUri.ToString() + " was not found on your system."
						);
                    }
					
				}
			}
			packDoc.Close();
			return result;
		}

		/// <summary>
		/// Function checks whether Document si Master/sub doc or simple doc
		/// </summary>
		/// <param name="listSubDocs">List of Documents</param>
		/// <returns> Message saying Liist of Docs are simple docs or not</returns>
		public static string CheckingSubDocs(ArrayList listSubDocs)
		{
			PackageRelationship relationship = null;
			String resultSubDoc = "simple";
			//TODO:
			for (int i = 0; i < listSubDocs.Count; i++)
			{
				string[] splt = listSubDocs[i].ToString().Split('|');
				Package pack;
				pack = Package.Open(splt[0].ToString(), FileMode.Open, FileAccess.ReadWrite);

				foreach (PackageRelationship searchRelation in pack.GetRelationshipsByType(WordRelationshipType))
				{
					relationship = searchRelation;
					break;
				}
				if (relationship == null) continue;

				Uri partUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				PackagePart mainPartxml = pack.GetPart(partUri);

				foreach (PackageRelationship searchRelation in mainPartxml.GetRelationships())
				{
					relationship = searchRelation;
					//checking whether Doc is simple or Master\sub Doc
					if (relationship.RelationshipType == WordSubdocumentRelationshipType)
					{
						if (relationship.TargetMode.ToString() == "External")
						{
							resultSubDoc = "complex";
						}
					}
				}
				pack.Close();
			}
			return resultSubDoc;
		}

		private static bool IsExternalSubdocumentRelationship(PackageRelationship relationship)
		{
			return relationship.RelationshipType == WordSubdocumentRelationshipType
			       && relationship.TargetMode.ToString() == "External";
		}

		private static bool IsMsWordSubdocumentRelationship(PackageRelationship relationship)
		{
			string targetUri = relationship.TargetUri.ToString();
			return Path.GetExtension(targetUri).Equals(".docx", StringComparison.InvariantCultureIgnoreCase)
				   || Path.GetExtension(targetUri).Equals(".doc", StringComparison.InvariantCultureIgnoreCase);
		}

		/// <summary>
		/// Return the system file path of a subdocument (without the file:// prefix), given a file path obtained from an input file 
		/// </summary>
		/// <param name="filePath">The file path referenced in the input (Master) file</param>
		/// <param name="inputPath">The path of the input (Master) file</param>
		/// <returns>The real file path if the file exists, or an empty string</returns>
		private static string GetRealFilePath(string filePath, string inputPath)
		{
			
			if (filePath.Contains("file") && filePath.Contains(LocalSettingsTemp))
			{
				filePath = filePath.Replace("file:///", "");
				int indx = filePath.LastIndexOf(LocalSettingsTemp);
				filePath = filePath.Substring(indx + LocalSettingsTemp.Length);
				filePath = Path.GetDirectoryName(inputPath) + "//" + filePath;
				if (File.Exists(filePath))
				{
					return filePath;
				}
			}
			else if (filePath.Contains("file"))
			{
				filePath = filePath.Replace("file:///", "");
				if (File.Exists(filePath))
				{
					return filePath;
				}
			}
			else
			{
				filePath = Path.GetDirectoryName(inputPath) + "\\" + filePath;
				if (File.Exists(filePath))
				{
					return filePath;
				}
			}
			return string.Empty;
		}
	}
}