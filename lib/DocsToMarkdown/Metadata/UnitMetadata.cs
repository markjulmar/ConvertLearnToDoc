using System.Text;
using HtmlAgilityPack;
using MSLearnCatalogAPI;

namespace Julmar.DocsToMarkdown.Metadata;

public class UnitMetadata : IDocsMetadata
{
    public ModuleUnit? CatalogMetadata { get; private set; }
    public string? Id { get; set; }
    public string? ZonePivotGroups { get; set; }
    public string? Author { get; set; }
    public string? MsAuthor { get; set; }
    public string? Title { get; set; }
    public string? Description { get; set; }
    public bool AzureSandbox { get; set; }
    public bool Sandbox { get; set; }
    public string? Keywords { get; set; }
    public string? GitUrl { get; set; }
    public string? Manager { get; set; }
    public string? MsCustom { get; set; }
    public string? MsDate { get; set; }
    public string? MsService { get; set; }
    public string? MsDevLang { get; set; }
    public string? MsTopic { get; set; }
    public string? PageKind { get; set; }
    public string? UpdatedAt { get; set; }
    private UnitMetadata()
    {
    }

    public static async Task<UnitMetadata> Create(HtmlDocument htmlDoc)
    {
        var metadata = new UnitMetadata();

        // Get the first H1 - this is the title.
        var title = htmlDoc.DocumentNode.SelectSingleNode("//h1");
        metadata.Title = title?.InnerText.Trim();

        var metadataTags = htmlDoc.DocumentNode.SelectNodes("//head/meta");
        foreach (var md in metadataTags.Where(md => md.Attributes.Contains("name")
                                                    && md.Attributes.Contains("content")))
        {
            string name = md.Attributes["name"].Value;
            string value = md.Attributes["content"].Value;

            switch (name)
            {
                case "uid":
                    metadata.Id = value;
                    break;
                case "zone_pivot_groups":
                    metadata.ZonePivotGroups = value;
                    break;
                case "author":
                    metadata.Author = value;
                    break;
                case "description":
                    metadata.Description = value;
                    break;
                case "original_content_git_url":
                    metadata.GitUrl = value;
                    break;
                case "manager":
                    metadata.Manager = value;
                    break;
                case "ms.author":
                    metadata.MsAuthor = value;
                    break;
                case "keywords":
                    metadata.Keywords = value;
                    break;
                case "ms.custom":
                    metadata.MsCustom = value;
                    break;
                case "ms.date":
                    metadata.MsDate = value;
                    break;
                case "ms.service":
                    metadata.MsService = value;
                    break;
                case "ms.devlang":
                    metadata.MsDevLang = value;
                    break;
                case "ms.topic":
                    metadata.MsTopic = value;
                    break;
                case "page_kind":
                    metadata.PageKind = value;
                    break;
                case "content_update_time":
                    metadata.UpdatedAt = value;
                    break;
                case "azure_sandbox":
                    metadata.AzureSandbox = value == "true";
                    break;
                case "sandbox":
                    metadata.Sandbox = value == "true";
                    break;
                default:
                    break;
            }
        }

        if (!string.IsNullOrEmpty(metadata.Id))
            await metadata.FillInFromLearnApi();
        
        return metadata;
    }

    private async Task FillInFromLearnApi()
    {
        var catalog = await CatalogApi.GetCatalogAsync();
        CatalogMetadata = catalog.Units.SingleOrDefault(unit => unit.Uid == Id);
    }

    public string YamlHeader
    {
        get
        {
            var sb = new StringBuilder("### YamlMime:ModuleUnit" + Environment.NewLine);
            if (Id != null) sb.AppendLine($"uid: {Id}");
            if (Title != null) sb.AppendLine($"title: {Title}");
            sb.AppendLine("metadata:");
            if (Title != null) sb.AppendLine($"  title: {Title}");
            if (Description != null) sb.AppendLine($"  description: {Description}");
            if (Keywords != null) sb.AppendLine($"  keywords: {Keywords}");
            if (MsTopic != null) sb.AppendLine($"  ms.topic: {MsTopic}");
            if (MsDate != null) sb.AppendLine($"  ms.date: {MsDate}");
            if (MsCustom != null) sb.AppendLine($"  ms.custom: {MsCustom}");
            if (Author != null) sb.AppendLine($"  author: {Author}");
            if (MsAuthor != null) sb.AppendLine($"  ms.author: {MsAuthor}");
            if (Manager != null) sb.AppendLine($"  manager: {Manager}");
            if (MsDevLang != null) sb.AppendLine($"  ms.devlang: {MsDevLang}");
            if (MsService != null) sb.AppendLine($"  ms.service: {MsService}");
            if (ZonePivotGroups != null) sb.AppendLine($"  zone_pivot_groups: {ZonePivotGroups}");
            if (AzureSandbox || Sandbox) sb.AppendLine("$  sandbox: true");

            if (CatalogMetadata != null)
            {
                if (CatalogMetadata.Duration > 0) sb.AppendLine($"durationInMinutes: {CatalogMetadata.Duration}");
            }

            return sb.ToString();
        }

    }
}