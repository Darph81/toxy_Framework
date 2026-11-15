using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
	/// <summary>
	/// Text extraction on top of the Open XML SDK model of a .docx file.
	/// </summary>
	internal static class WordprocessingText
	{
		/// <summary>
		/// Enumerates the paragraphs and tables directly contained in <paramref name="container"/> in document order,
		/// including the ones wrapped in content controls.
		/// </summary>
		public static IEnumerable<OpenXmlElement> GetBlocks(OpenXmlElement container)
		{
			foreach (OpenXmlElement child in container.ChildElements)
			{
				switch (child)
				{
					case Paragraph _:
					case Table _:
						yield return child;
						break;
					case SdtBlock sdt:
						SdtContentBlock content = sdt.SdtContentBlock;
						if (content != null)
						{
							foreach (OpenXmlElement block in GetBlocks(content))
							{
								yield return block;
							}
						}
						break;
					case CustomXmlBlock custom:
						foreach (OpenXmlElement block in GetBlocks(custom))
						{
							yield return block;
						}
						break;
				}
			}
		}

		private static IEnumerable<Paragraph> GetParagraphs(OpenXmlElement container)
		{
			return GetBlocks(container).OfType<Paragraph>();
		}

		/// <summary>The paragraphs directly contained in the cells of <paramref name="table"/>.</summary>
		public static IEnumerable<Paragraph> GetCellParagraphs(Table table)
		{
			return GetRows(table).SelectMany(row => row.Elements<TableCell>()).SelectMany(GetParagraphs);
		}

		private static IEnumerable<TableRow> GetRows(Table table)
		{
			return table.Elements<TableRow>();
		}

		/// <summary>True if the option is set to a true value (1, on, true) in the properties of the context.</summary>
		public static bool IsOptionEnabled(ParserContext context, string name)
		{
			return context.Properties.ContainsKey(name) && Utility.IsTrue(context.Properties[name]);
		}

		/// <summary>The text of every header of the document.</summary>
		public static IEnumerable<string> GetHeaderTexts(MainDocumentPart mainPart)
		{
			return mainPart.HeaderParts.Select(part => GetBlocksText(part.Header));
		}

		/// <summary>The text of every footer of the document.</summary>
		public static IEnumerable<string> GetFooterTexts(MainDocumentPart mainPart)
		{
			return mainPart.FooterParts.Select(part => GetBlocksText(part.Footer));
		}

		public static string GetStyleId(Paragraph paragraph)
		{
			return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		}

		/// <summary>The text of the runs of a paragraph. Tabs and line, page and column breaks are rendered as '\t' and '\n'.</summary>
		public static string GetText(Paragraph paragraph)
		{
			StringBuilder builder = new StringBuilder();
			foreach (Run run in paragraph.Descendants<Run>())
			{
				// runs of paragraphs nested in the paragraph (e.g. text boxes) belong to the nested paragraph
				if (run.Ancestors<Paragraph>().First() != paragraph)
				{
					continue;
				}
				foreach (OpenXmlElement element in run.ChildElements)
				{
					switch (element)
					{
						case Text text: builder.Append(text.Text); break;
						case TabChar _: builder.Append('\t'); break;
						case Break _:
						case CarriageReturn _: builder.Append('\n'); break;
						case NoBreakHyphen _: builder.Append('-'); break;
					}
				}
			}
			return builder.ToString();
		}

		/// <summary>The text of a table: one line per row, cells separated by tabs.</summary>
		public static string GetText(Table table)
		{
			return string.Join("\n", GetRows(table).Select(row => string.Join("\t",
				row.Elements<TableCell>().Select(cell => string.Join("\n", GetParagraphs(cell).Select(GetText))))));
		}

		/// <summary>The text of a header, footer or other container of block elements, one line per block.</summary>
		public static string GetBlocksText(OpenXmlElement container)
		{
			return string.Join("\n", GetBlocks(container).Select(block =>
			{
				switch (block)
				{
					case Paragraph paragraph: return GetText(paragraph);
					case Table table: return GetText(table);
					default: return string.Empty;
				}
			}));
		}
	}
}
