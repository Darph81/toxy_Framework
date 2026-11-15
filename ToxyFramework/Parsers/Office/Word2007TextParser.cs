using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	public class Word2007TextParser : BaseTextParser
	{
		public Word2007TextParser(ParserContext context) : base(context)
		{ }

		internal override void ValidateContext()
		{
			base.ValidateContext();
			Utility.ThrowIfProtected(Context);
		}

		internal override string ParseText(Stream stream)
		{
			bool extractHeader = WordprocessingText.IsOptionEnabled(Context, "ExtractHeader");
			bool extractFooter = WordprocessingText.IsOptionEnabled(Context, "ExtractFooter");

			StringBuilder sb = new StringBuilder();
			using (WordprocessingDocument worddoc = WordprocessingDocument.Open(stream, false))
			{
				MainDocumentPart mainPart = worddoc.MainDocumentPart;
				if (extractHeader && mainPart != null)
				{
					foreach (string header in WordprocessingText.GetHeaderTexts(mainPart))
					{
						sb.Append("[Header] ").AppendLine(header);
					}
				}

				Body body = mainPart?.Document?.Body;
				if (body != null)
				{
					foreach (var block in WordprocessingText.GetBlocks(body))
					{
						if (block is Paragraph paragraph)
						{
							sb.AppendLine(WordprocessingText.GetText(paragraph));
						}
						else if (block is Table table)
						{
							sb.AppendLine(WordprocessingText.GetText(table));
						}
					}
				}

				if (extractFooter && mainPart != null)
				{
					foreach (string footer in WordprocessingText.GetFooterTexts(mainPart))
					{
						sb.Append("[Footer] ").AppendLine(footer);
					}
				}
			}
			return sb.ToString();
		}
	}
}
