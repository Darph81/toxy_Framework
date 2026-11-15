using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace Toxy.Parsers
{
	public class OOXMLMetadataParser : IMetadataParser
	{
		private const string ExtendedPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

		public OOXMLMetadataParser(ParserContext context)
		{
			Context = context;
		}

		public ParserContext Context { get; set; }

		public ToxyMetadata Parse()
		{
			Utility.ValidateContext(Context);

			ToxyMetadata metadata = new ToxyMetadata();
			Package pack = null;
			try
			{
				if (Context.IsStreamContext)
				{
					pack = Package.Open(Context.Stream, FileMode.Open, FileAccess.Read);
				}
				else
				{
					pack = Package.Open(Context.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
				}

				AddCoreProperty(metadata, pack.PackageProperties);
				AddExtendedProperties(metadata, ReadExtendedProperties(pack));
			}
			finally
			{
				pack?.Close();
			}
			return metadata;
		}

		/// <summary>
		/// Adds the first core property which is set. Only one of them is reported, this is the behavior this parser always had.
		/// </summary>
		private static void AddCoreProperty(ToxyMetadata metadata, PackageProperties props)
		{
			bool Add(string name, object value)
			{
				if (value == null)
				{
					return false;
				}
				metadata.Add(name, value);
				return true;
			}

			_ = Add("Title", props.Title) || Add("Identifier", props.Identifier) || Add("Keywords", props.Keywords)
				|| Add("Revision", props.Revision) || Add("Subject", props.Subject) || Add("Modified", props.Modified)
				|| Add("LastPrinted", props.LastPrinted) || Add("Created", props.Created) || Add("Creator", props.Creator)
				|| Add("Description", props.Description);
		}

		private static void AddExtendedProperties(ToxyMetadata metadata, IReadOnlyDictionary<string, string> props)
		{
			foreach (string name in new[] { "Application", "AppVersion", "Company", "Manager" })
			{
				if (props.TryGetValue(name, out string text))
				{
					metadata.Add(name, text);
				}
			}

			// counters are only reported if they are greater than zero
			foreach (string name in new[] { "Characters", "CharactersWithSpaces", "Lines", "Notes", "Pages", "Paragraphs", "Words", "TotalTime" })
			{
				if (props.TryGetValue(name, out string text)
					&& long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number)
					&& number > 0)
				{
					metadata.Add(name, number);
				}
			}
		}

		/// <summary>
		/// Reads the text of the extended (application) properties, i.e. the part /docProps/app.xml, by their local name.
		/// </summary>
		private static IReadOnlyDictionary<string, string> ReadExtendedProperties(Package package)
		{
			var properties = new Dictionary<string, string>();
			PackageRelationship relationship = package.GetRelationshipsByType(ExtendedPropertiesRelationship).FirstOrDefault();
			if (relationship == null)
			{
				return properties;
			}

			Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
			if (!package.PartExists(partUri))
			{
				return properties;
			}

			using (var stream = package.GetPart(partUri).GetStream())
			{
				foreach (XElement element in XDocument.Load(stream).Root.Elements())
				{
					properties[element.Name.LocalName] = element.Value;
				}
			}
			return properties;
		}
	}
}
