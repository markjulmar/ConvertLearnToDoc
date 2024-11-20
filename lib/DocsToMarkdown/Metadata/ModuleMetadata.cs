using System.Text;
using HtmlAgilityPack;
using MSLearnCatalogAPI;

namespace Julmar.DocsToMarkdown.Metadata;

public sealed class ModuleMetadata : BaseMetadata
{
    private Module? _catalogMetadata;
    public override string? PageKind => "module";

    private ModuleMetadata(HtmlDocument document) : base(document)
    {
    }

    public static async Task<IDocsMetadata> Create(HtmlDocument document)
    {
        var metadata = new ModuleMetadata(document);
        metadata.FillInFromPageData(document);
        await metadata.FillInFromLearnApi();
        return metadata;
    }

    private async Task FillInFromLearnApi()
    {
        var id = LoadedValues.GetValueOrDefault("uid");
        if (string.IsNullOrWhiteSpace(id))
            return;
        
        var modules = await CatalogApi.GetModulesAsync(filter: new CatalogFilter()
        {
            Uids = [id]
        });
        
        if (modules.Count == 1)
            _catalogMetadata = modules[0];
    }

    private void FillInFromPageData(HtmlDocument htmlDocument)
    {
        var converter = new HtmlConverter(new Uri("https://localhost"), htmlDocument, null);

        var allDivs = htmlDocument.DocumentNode
            .SelectNodes("//div").ToList();

        var summaryNode = allDivs.SingleOrDefault(n => n.GetClasses().Contains("learn-summary"));
        if (summaryNode != null)
        {
            LoadedValues.Add("summary", converter.Convert(summaryNode).TrimEnd('\r', '\n', ' '));
        }

        var abstractNode = allDivs.SingleOrDefault(n => n.GetClasses().Contains("abstract"));
        if (abstractNode != null)
        {
            LoadedValues.Add("abstract", converter.Convert(abstractNode).TrimEnd('\r', '\n', ' '));
        }

        var prerequisitesNode = allDivs.SingleOrDefault(n => n.GetClasses().Contains("prerequisites"));
        if (prerequisitesNode != null)
        {
            prerequisitesNode.Element("h2")?.Remove(); // remove any header
            LoadedValues.Add("prerequisites", converter.Convert(prerequisitesNode).TrimEnd('\r', '\n', ' '));
        }
    }

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
    
    public override string YamlHeader
    {
        get
        {
            var sb = new StringBuilder("### YamlMime:Module" + Environment.NewLine);
            var id = LoadedValues.GetValueOrDefault("uid");
            if (!string.IsNullOrWhiteSpace(id))
                sb.AppendLine($"uid: {id}");

            sb.AppendLine("metadata:");
            foreach (var key in MetadataNames)
            {
                if (LoadedValues.TryGetValue(key, out var value)
                    && !string.IsNullOrWhiteSpace(value))
                    sb.AppendLine($"  {key}: {EscapeYaml(value)}");
            }
            
            var msCollection = LoadedValues.GetValueOrDefault("ms.collection");
            if (!string.IsNullOrEmpty(msCollection))
            {
                sb.AppendLine("  ms.collection:");
                foreach (var item in msCollection.Split(','))
                {
                    sb.AppendLine($"    - {item.Trim()}");
                }
            }

            var title = LoadedValues.GetValueOrDefault("title");
            if (!string.IsNullOrWhiteSpace(title))
                sb.AppendLine($"title: {EscapeYaml(title)}");

            var summary = LoadedValues.GetValueOrDefault("summary")
                          ?? _catalogMetadata?.Summary;
            if (!string.IsNullOrEmpty(summary))
                sb.AppendLine($"summary: {EscapeYaml(summary)}");
            
            var @abstract = LoadedValues.GetValueOrDefault("abstract");
            if (!string.IsNullOrEmpty(@abstract))
                sb.AppendLine($"abstract: {EscapeYaml(@abstract)}");
                          
            var prerequisites = LoadedValues.GetValueOrDefault("prerequisites");
            if (!string.IsNullOrEmpty(prerequisites))
                sb.AppendLine($"prerequisites: {EscapeYaml(prerequisites)}");
            
            if (_catalogMetadata != null)
            {
                const string prefix = @"https://learn.microsoft.com/en-us/";
                sb.AppendLine($"iconUrl: /{_catalogMetadata.IconUrl[prefix.Length..]}");
                if (_catalogMetadata.Levels.Count > 0)
                {
                    sb.AppendLine("levels:");
                    foreach (var level in _catalogMetadata.Levels)
                        sb.AppendLine($"- {level}");
                }

                if (_catalogMetadata.Roles.Count > 0)
                {
                    sb.AppendLine("roles:");
                    foreach (var role in _catalogMetadata.Roles)
                        sb.AppendLine($"- {role}");
                }
                
                if (_catalogMetadata.Products.Count > 0)
                {
                    sb.AppendLine("products:");
                    foreach (var product in _catalogMetadata.Products)
                        sb.AppendLine($"- {product}");
                }

                if (_catalogMetadata.Subjects.Count > 0)
                {
                    sb.AppendLine("subjects:");
                    foreach (var subject in _catalogMetadata.Subjects)
                        sb.AppendLine($"- {subject}");
                }

                if (_catalogMetadata.Units.Count > 0)
                {
                    sb.AppendLine("units:");
                    foreach (var unit in _catalogMetadata.Units)
                        sb.AppendLine($"- {unit}");
                }

                if (!string.IsNullOrWhiteSpace(id))
                    sb.AppendLine($"badge:").AppendLine($"  uid: {id}.badge");
            }

            return sb.ToString();
        }
    }
}