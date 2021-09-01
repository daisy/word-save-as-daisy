using System;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion.PageStylesValidation
{
	public class PageStylesValidator
	{
		protected PageStyle _lastStyle = PageStyle.Empty;

		public ValidatoinResult ValidateParagraph(IList<PageStyle> paragraphStyles)
		{
			ValidatoinResult result = new ValidatoinResult();
			if (paragraphStyles.Count != 0)
			{

				if (paragraphStyles.Count > 1)
				{
					result = new ValidatoinResult("More than one page style applied to paragraph, please applay only one style from Frontmatter (DAISY), Bodymatter (DAISY) and Rearmatter (DAISY)");
					_lastStyle = paragraphStyles[paragraphStyles.Count - 1];
				}

				else
				{
					PageStyle currentStyle = paragraphStyles[0];
					switch (_lastStyle)
					{
						case PageStyle.Empty:
							if (currentStyle == PageStyle.Rearmatter)
							{
								result = new ValidatoinResult("Rearmatter (Daisy) can be applied only after Bodymetter (DAISY) style");
							}
							break;
						case PageStyle.Frontmatter:
							if (currentStyle == PageStyle.Frontmatter)
							{
								result = new ValidatoinResult("Frontmatter (DAISY) style should be applied only one time");
							}
							else if (currentStyle == PageStyle.Rearmatter)
							{
								result = new ValidatoinResult("Rearmatter (DAISY) style can be applied only after Bodymatter (DAISY) style");
							}
							break;
						case PageStyle.Bodymatter:
							if (currentStyle == PageStyle.Frontmatter)
							{
								result = new ValidatoinResult("Frontmatter (DAISY) style can not be applied after Bodymatter (DAISY) style");
							}
							else if (currentStyle == PageStyle.Bodymatter)
							{
								result = new ValidatoinResult("Bodymatter (DAISY) style should be applied only one time");
							}
							break;
						case PageStyle.Rearmatter:
							result = new ValidatoinResult("Frontmatter (DAISY), Bodymatter (DAISY) and Rearmatter (DAISY) styles can not be applied after Rearmatter (DAISY) style");
							break;
						default:
							throw new ArgumentOutOfRangeException();
					}
					_lastStyle = currentStyle;
				}
			}

			return result;
		}

		public virtual ValidatoinResult ValidateLastStyle()
		{
			if (_lastStyle == PageStyle.Frontmatter)
				return new ValidatoinResult("Frontmatter (DAISY) can not be applied without Bodymatter (DAISY) style ");
			return new ValidatoinResult();
		}

		public static PageStyle GetPageStyle(string style)
		{
			if (style.StartsWith("Frontmatter"))
				return PageStyle.Frontmatter;

			if (style.StartsWith("Bodymatter"))
				return PageStyle.Bodymatter;

			if (style.StartsWith("Rearmatter"))
				return PageStyle.Rearmatter;

			return PageStyle.Empty;
		}
	}
}