using System;
using System.Collections;
using System.IO;
using System.IO.Packaging;

namespace Sonata.DaisyConverter.DaisyConverterLib.Converters
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
				if (!IsExternalSubdocumentRelationship(relationship)) 
					continue;
				
				if (IsMsWordSubdocumentRelationship(relationship))
					result.SubdocumentsCount++;

				String fileName = GetFileName(relationship.TargetUri.ToString(), originalInputPath);

				if (string.IsNullOrEmpty(fileName))
					continue;

				if (IsMsWordSubdocumentRelationship(relationship))
				{
					result.Subdocuments.Add(new SubdocumentInfo(fileName, relationship.Id));
				}
				else
				{
					result.NotTranslatedSubdocuments.Add(new SubdocumentInfo(fileName, relationship.Id));
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

		private static string GetFileName(string fileName, string inputPath)
		{
			if (fileName.Contains("file") && fileName.Contains(LocalSettingsTemp))
			{
				fileName = fileName.Replace("file:///", "");
				int indx = fileName.LastIndexOf(LocalSettingsTemp);
				fileName = fileName.Substring(indx + LocalSettingsTemp.Length);
				fileName = Path.GetDirectoryName(inputPath) + "//" + fileName;
				if (File.Exists(fileName))
				{
					return fileName;
				}
			}
			else if (fileName.Contains("file"))
			{
				fileName = fileName.Replace("file:///", "");
				if (File.Exists(fileName))
				{
					return fileName;
				}
			}
			else
			{
				fileName = Path.GetDirectoryName(inputPath) + "\\" + Path.GetFileName(fileName);
				if (File.Exists(fileName))
				{
					return fileName;
				}
			}
			return string.Empty;
		}
	}
}