using System.Collections;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion
{
	public class SubdocumentsList
	{
		public SubdocumentsList()
		{
			Subdocuments = new List<SubdocumentInfo>();
			NotTranslatedSubdocuments = new List<SubdocumentInfo>();
			Errors = new List<string>();
		}

		public bool Empty { 
			get => Subdocuments.Count == 0;
		}

		public int SubdocumentsCount { get; set; }
		public List<SubdocumentInfo> Subdocuments { get; set; }
		public List<SubdocumentInfo> NotTranslatedSubdocuments { get; set; }

		public List<string> Errors { get; set; }

		public ArrayList GetSubdocumentsNames()
		{
			return GetFileNames(Subdocuments);
		}

		public ArrayList GetNotTraslatedSubdocumentsNames()
		{
			return GetFileNames(NotTranslatedSubdocuments);
		}

		public ArrayList GetSubdocumentsNameWithRelationship()
		{
			return GetFileNamesWithRelationship(Subdocuments);
		}

		private ArrayList GetFileNames(List<SubdocumentInfo> subdocuments)
		{
			ArrayList result = new ArrayList(subdocuments.Count);
			foreach (var subdocumentInfo in subdocuments)
			{
				result.Add(subdocumentInfo.FileName);
			}
			return result;
		}

		private ArrayList GetFileNamesWithRelationship(List<SubdocumentInfo> subdocuments)
		{
			ArrayList result = new ArrayList(subdocuments.Count);
			foreach (var subdocumentInfo in subdocuments)
			{
				result.Add(subdocumentInfo.FileNameWithRelationship);
			}
			return result;
		}
	}
}