using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Metadata;

public abstract class BaseMetadata : IDocsMetadata
{
    private static readonly string[] MetadataNames =
    [
        "title",
        "description",
        "keywords",
        "ms.topic",
        "ms.date",
        "ms.custom",
        "author",
        "ms.author",
        "manager",
        "ms.service",
        "ms.devlang",
        "zone_pivot_groups",
    ];

    protected readonly Dictionary<string, string> LoadedValues = new();

    public abstract string? PageKind { get; }

    public virtual string YamlHeader
    {
        get
        {
            var sb = new StringBuilder();
            foreach (var key in MetadataNames)
            {
                if (Values.TryGetValue(key, out var value)
                    && !string.IsNullOrWhiteSpace(value))
                {
                    sb.AppendLine($"{key}: {EscapeYaml(value)}");
                }
            }

            var msCollection = LoadedValues.GetValueOrDefault("ms.collection");
            if (!string.IsNullOrEmpty(msCollection))
            {
                sb.AppendLine("ms.collection:");
                foreach (var item in msCollection.Split(','))
                {
                    sb.AppendLine($"  - {item.Trim()}");
                }
            }
            
            return sb.ToString();
        }
    }

    protected string EscapeYaml(string value)
    {
        if (value.Contains('\n') || value.TrimStart().StartsWith('-'))
            return $" |\n{string.Join('\n', value.Split('\n')
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => $"  {line}"))}";
        if (value.Contains(':'))
            return $"\"{value}\"";
        return value;
    }

    public IReadOnlyDictionary<string, string> Values => LoadedValues;
    
    protected BaseMetadata(HtmlDocument htmlDocument)
    {
        var metadataTags = htmlDocument.DocumentNode.SelectNodes("//head/meta");
        foreach (var md in metadataTags.Where(md => 
                     md.Attributes.Contains("name")
                     && md.Attributes.Contains("content")))
        {
            var name = md.Attributes["name"].Value;
            var value = md.Attributes["content"].Value;

            // Special case for date
            if (name == "ms.date")
                value = DateTime.TryParse(value, out var date) ? date.ToString("MM/dd/yyyy") : value;
            
            LoadedValues[name] = value;
        }
        
        // Get the first H1 - this is the title.
        // Prefer this over the one in metadata.
        var title = htmlDocument.DocumentNode.SelectSingleNode("//h1");
        if (title != null)
            LoadedValues["title"] = title.InnerText.Trim();
    }
}