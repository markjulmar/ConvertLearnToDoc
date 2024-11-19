using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;
using MSLearnCatalogAPI;

namespace Julmar.DocsToMarkdown.Metadata;

public class ModuleMetadata : IDocsMetadata
{
    public Module? CatalogMetadata { get; private set; }
    public string? Id { get; set; }
    public string? ZonePivotGroups { get; set; }
    public string? Author { get; set; }
    public string? MsAuthor { get; set; }
    public string? Title { get; set; }
    public string? Description { get; set; }
    public string? Keywords { get; set; }
    public string? GitUrl { get; set; }
    public string? Manager { get; set; }
    public string? MsCustom { get; set; }
    public string? MsDate { get; set; }
    public string? MsService { get; set; }
    public string? MsCollection { get; set;}
    public string? MsDevLang { get; set; }
    public string? MsTopic { get; set; }
    public string? PageKind { get; set; }
    public string? UpdatedAt { get; set; }

    private ModuleMetadata()
    {
    }
    
    public static async Task<IDocsMetadata> Create(HtmlDocument htmlDoc)
    {
        var metadata = new ModuleMetadata();

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
                case "ms.collection":
                    metadata.MsCollection = value;
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
        Debug.Assert(!string.IsNullOrEmpty(Id));
        var modules = await CatalogApi.GetModulesAsync(filter: new CatalogFilter()
        {
            Uids = [Id!]
        });
        
        if (modules.Count == 1)
            CatalogMetadata = modules[0];
    }

    public string YamlHeader
    {
        get
        {
            var sb = new StringBuilder("### YamlMime:Module" + Environment.NewLine);
            if (Id != null) sb.AppendLine($"uid: {Id}");
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
            if (MsCollection != null) sb.AppendLine($"  ms.collection:{Environment.NewLine}    - {MsCollection}");

            if (Title != null) sb.AppendLine($"title: {Title}");
            if (CatalogMetadata != null)
            {
                const string prefix = @"https://learn.microsoft.com/en-us/";
                sb.AppendLine($"summary: {CatalogMetadata.Summary}");
                sb.AppendLine("iconUrl: " + CatalogMetadata.IconUrl[prefix.Length..]);
                if (CatalogMetadata.Levels.Count > 0)
                {
                    sb.AppendLine("levels:");
                    foreach (var level in CatalogMetadata.Levels)
                        sb.AppendLine($"- {level}");
                }

                if (CatalogMetadata.Roles.Count > 0)
                {
                    sb.AppendLine("roles:");
                    foreach (var role in CatalogMetadata.Roles)
                        sb.AppendLine($"- {role}");
                }
                
                if (CatalogMetadata.Products.Count > 0)
                {
                    sb.AppendLine("products:");
                    foreach (var product in CatalogMetadata.Products)
                        sb.AppendLine($"- {product}");
                }

                if (CatalogMetadata.Subjects.Count > 0)
                {
                    sb.AppendLine("subjects:");
                    foreach (var subject in CatalogMetadata.Subjects)
                        sb.AppendLine($"- {subject}");
                }

                if (CatalogMetadata.Units.Count > 0)
                {
                    sb.AppendLine("units:");
                    foreach (var unit in CatalogMetadata.Units)
                        sb.AppendLine($"- {unit}");
                }

                sb.AppendLine($"badge: {Id}.badge");
            }

            return sb.ToString();
        }
    }
}