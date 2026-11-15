using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
	public class Word2007DocumentParser : IDocumentParser
	{
		public ParserContext Context { get; set; }

		public Word2007DocumentParser(ParserContext context)
		{
			Context = context;
		}

		public ToxyDocument Parse()
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);
			bool extractHeader = WordprocessingText.IsOptionEnabled(Context, "ExtractHeader");
			bool extractFooter = WordprocessingText.IsOptionEnabled(Context, "ExtractFooter");

			ToxyDocument rdoc = new ToxyDocument();
			Stream stream = Utility.GetStream(Context);
			try
			{
				using (WordprocessingDocument worddoc = WordprocessingDocument.Open(stream, false))
				{
					MainDocumentPart mainPart = worddoc.MainDocumentPart;
					if (mainPart == null)
					{
						return rdoc;
					}

					if (extractHeader)
					{
						rdoc.Header = JoinLines(WordprocessingText.GetHeaderTexts(mainPart));
					}
					if (extractFooter)
					{
						rdoc.Footer = JoinLines(WordprocessingText.GetFooterTexts(mainPart));
					}

					Body body = mainPart.Document?.Body;
					if (body != null)
					{
						// body paragraphs first, followed by the paragraphs of the tables
						List<Paragraph> tableParagraphs = new List<Paragraph>();
						foreach (OpenXmlElement block in WordprocessingText.GetBlocks(body))
						{
							if (block is Paragraph para)
							{
								rdoc.Paragraphs.Add(ToToxyParagraph(para));
							}
							else if (block is Table table)
							{
								tableParagraphs.AddRange(WordprocessingText.GetCellParagraphs(table));
							}
						}
						foreach (Paragraph cellPara in tableParagraphs)
						{
							rdoc.Paragraphs.Add(ToToxyParagraph(cellPara));
						}
					}
				}
			}
			finally
			{
				if (!Context.IsStreamContext)
				{
					stream.Dispose();
				}
			}
			return rdoc;
		}

		private static string JoinLines(IEnumerable<string> lines)
		{
			StringBuilder sb = new StringBuilder();
			foreach (string line in lines)
			{
				sb.AppendLine(line);
			}
			return sb.ToString();
		}

		private static ToxyParagraph ToToxyParagraph(Paragraph paragraph)
		{
			ToxyParagraph p = new ToxyParagraph();
			p.Text = WordprocessingText.GetText(paragraph);
			p.StyleID = WordprocessingText.GetStyleId(paragraph);
			return p;
		}
	}
}
