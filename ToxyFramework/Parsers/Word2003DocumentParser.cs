using NPOI.XWPF.UserModel;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class Word2003DocumentParser : IDocumentParser
    {
        public Word2003DocumentParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyDocument Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            bool extractHeader = false;
            if (Context.Properties.ContainsKey("ExtractHeader"))
            {
                extractHeader = Utility.IsTrue(Context.Properties["ExtractHeader"]);
            }
            bool extractFooter = false;
            if (Context.Properties.ContainsKey("ExtractFooter"))
            {
                extractFooter = Utility.IsTrue(Context.Properties["ExtractFooter"]);
            }

            ToxyDocument rdoc = new ToxyDocument();


            using (FileStream stream = File.OpenRead(Context.Path))
            {
                XWPFDocument wordDoc = new XWPFDocument(stream);

                if (extractHeader && wordDoc.HeaderList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var header in wordDoc.HeaderList)
                    {
                        sb.AppendLine(header.Text);
                    }
                    rdoc.Header = sb.ToString();
                }

                if (extractFooter && wordDoc.FooterList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var footer in wordDoc.FooterList)
                    {
                        sb.AppendLine(footer.Text);
                    }
                    rdoc.Footer = sb.ToString();
                }

                foreach (XWPFParagraph para in wordDoc.Paragraphs)
                {
                    string text = para.Text;
                    ToxyParagraph p = new ToxyParagraph();
                    p.Text = text;
                    p.StyleID = para.Style;

                    rdoc.Paragraphs.Add(p);
                }
            }
            return rdoc;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}
