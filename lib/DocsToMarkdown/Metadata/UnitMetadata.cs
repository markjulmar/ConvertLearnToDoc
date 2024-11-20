using System.Text;
using HtmlAgilityPack;
using MSLearnCatalogAPI;

namespace Julmar.DocsToMarkdown.Metadata;

public class UnitMetadata : BaseMetadata
{
    private ModuleUnit? _catalogMetadata;

    private UnitMetadata(HtmlDocument document) : base(document)
    {
    }

    public static async Task<UnitMetadata> Create(HtmlDocument htmlDoc)
    {
        var metadata = new UnitMetadata(htmlDoc);
        await metadata.FillInFromLearnApi();
        return metadata;
    }

    private async Task FillInFromLearnApi()
    {
        var id = LoadedValues.GetValueOrDefault("uid");
        if (string.IsNullOrWhiteSpace(id))
            return;
        
        var catalog = await CatalogApi.GetCatalogAsync();
        _catalogMetadata = catalog.Units.SingleOrDefault(unit => unit.Uid == id);
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

    private readonly Dictionary<string, string> _additionalKeys = new()
    {
        { "azure_sandbox", "azureSandbox" },
        { "sandbox", "sandbox" },
        { "lab-id", "labId" },
        { "labModal", "labModal" }
    };
    
    public override string? PageKind => "unit";

    public override string YamlHeader
    {
        get
        {
            var sb = new StringBuilder("### YamlMime:ModuleUnit" + Environment.NewLine);
            var id = LoadedValues.GetValueOrDefault("uid");
            if (!string.IsNullOrWhiteSpace(id))
                sb.AppendLine($"uid: {id}");
            var title = LoadedValues.GetValueOrDefault("title");
            if (!string.IsNullOrWhiteSpace(title))
                sb.AppendLine($"title: {EscapeYaml(title)}");
            sb.AppendLine("metadata:");
            foreach (var name in MetadataNames)
            {
                var value = LoadedValues.GetValueOrDefault(name);
                if (!string.IsNullOrWhiteSpace(value))
                    sb.AppendLine($"  {name}: {EscapeYaml(value)}");
            }

            foreach (var kvp in _additionalKeys)
            {
                var value = LoadedValues.GetValueOrDefault(kvp.Key);
                if (!string.IsNullOrWhiteSpace(value))
                    sb.AppendLine($"{kvp.Value}: {EscapeYaml(value)}");
            }
            
            if (_catalogMetadata is { Duration: > 0 }) 
                sb.AppendLine($"durationInMinutes: {_catalogMetadata.Duration}");

            return sb.ToString();
        }
    }
}