namespace Daisy.SaveAsDAISY.Conversion
{
	public class SubdocumentInfo
	{
		public SubdocumentInfo(string fileName, string relationshipId)
		{
			FileName = fileName;
			RelationshipId = relationshipId;
		}

		public string FileName { get; private set; }
		public string RelationshipId { get; private set; }
		public string FileNameWithRelationship { get { return FileName + "|" + RelationshipId; } }
	}
}