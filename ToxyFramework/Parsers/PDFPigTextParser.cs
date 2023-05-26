using System.IO;
using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Toxy.Parsers
{
    public class PDFPigTextParser : ITextParser
    {
        public ParserContext Context { get; set; }

        public PDFPigTextParser(ParserContext context)
        {
            this.Context = context;
        }

        public string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            StringBuilder sb = new StringBuilder();
            using (var pdf = PdfDocument.Open(Context.Path))
            {
                foreach (var page in pdf.GetPages())
                {
                    string text = ContentOrderTextExtractor.GetText(page);
                    if (string.IsNullOrEmpty(text))
                        text = string.Join(" ", page.GetWords());
                    if (string.IsNullOrEmpty(text))
                        text = page.Text;

                    sb.AppendLine(text);
                }
            }

            return sb.ToString();
        }
    }
}
