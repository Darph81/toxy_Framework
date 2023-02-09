using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Toxy;

namespace ToxyFramework.Parsers
{
    /// <summary>
    /// Parse Text content from ODS, ODT, ODP files in plain content
    /// </summary>
    public class ODSTextParser : ITextParser
    {
        private const string ContentFileName = "content.xml";

        public ParserContext Context
        {
            get;
            set;
        }

        public ODSTextParser(ParserContext context)
        {
            this.Context = context;
        }

        public string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                using (var zipArchive = new ZipArchive(stream))
                {
                    var contentEntry = zipArchive.Entries.SingleOrDefault(x => x.Name == ContentFileName);

                    if (contentEntry == null)
                        throw new InvalidOperationException("Can not find content.xml in file");

                    using (var contentEntryStream = contentEntry.Open())
                    {
                        var document = XDocument.Load(contentEntryStream);
                        sb.AppendLine(document.Root?.Value);
                    }
                }
            }

            return sb.ToString();
        }
    }
}
