using System;
using System.Collections.Generic;
using System.Xml;
using Daisy.SaveAsDAISY.Conversion.PageStylesValidation;
using Daisy.SaveAsDAISY.Forms;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	/// Provides functionality for checking sequence of Frontmatter (DAISY), Bodymatter (DAISY) and Rearmatter (DAISY) styles in word document
	/// </summary>
	public class OoxPageStylesValidator : PageStylesValidator
	{
		private readonly ProgressValidation _progressValidation;

		public OoxPageStylesValidator(ProgressValidation progressValidation)
		{
			_progressValidation = progressValidation;
		}

		public void ValidateNode(XmlNode node)
		{
			var pageStyles = GetPageStyleNodes(node);
			IList<PageStyle> styles = new List<PageStyle>();
			foreach (XmlNode pageStyle in pageStyles)
			{
				styles.Add(GetPageStyle(pageStyle));
			}

			var result = ValidateParagraph(styles);

			if (!result.IsValid)
				_progressValidation.AddValidationError(string.Format(" :{0}|PageStyles error {1}", result.ErrorMessage,
				                                                     _progressValidation.GenerateId()));

		}

		public override ValidatoinResult ValidateLastStyle()
		{
			var result = base.ValidateLastStyle();
			
			if(!result.IsValid)
				_progressValidation.AddValidationError(string.Format(" :{0}|PageStyles error {1}", result.ErrorMessage,
																					 _progressValidation.GenerateId()));
			return result;

		}

		private List<XmlNode> GetPageStyleNodes(XmlNode node)
		{
			List<XmlNode> styleNodes = new List<XmlNode>();
			if (node == null) return styleNodes;

			if ( (node.Name =="w:pStyle"  || node.Name == "w:rStyle") &&
					GetPageStyle(node) != PageStyle.Empty)
			{
				styleNodes.Add(node);
			}

			foreach (XmlNode childNode in node.ChildNodes)
				styleNodes.AddRange(GetPageStyleNodes(childNode));

			return styleNodes;
		}

		private PageStyle GetPageStyle(XmlNode styleNode)
		{
			if (styleNode == null) throw new ArgumentNullException("styleNode");

			if (styleNode.Attributes["w:val"] == null)
				return PageStyle.Empty;

			return GetPageStyle(styleNode.Attributes["w:val"].Value);
		}
	}
}