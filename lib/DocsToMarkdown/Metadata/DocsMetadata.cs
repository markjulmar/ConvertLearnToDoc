using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Metadata;

public class DocsMetadata : IDocsMetadata
{
    private DocsMetadata()
    {
    }
    
    public string YamlHeader
    {
        get
        {
            var sb = new StringBuilder();
            if (Title != null) sb.AppendLine($"title: {Title}");
            if (Description != null) sb.AppendLine($"description: {Description}");
            if (Keywords != null) sb.AppendLine($"keywords: {Keywords}");
            if (MsTopic != null) sb.AppendLine($"ms.topic: {MsTopic}");
            if (MsDate != null) sb.AppendLine($"ms.date: {MsDate}");
            if (MsCustom != null) sb.AppendLine($"ms.custom: {MsCustom}");
            if (Author != null) sb.AppendLine($"author: {Author}");
            if (MsAuthor != null) sb.AppendLine($"ms.author: {MsAuthor}");
            if (AssetId != null) sb.AppendLine($"ms.assetid: {AssetId}");
            if (Manager != null) sb.AppendLine($"manager: {Manager}");
            if (MsDevLang != null) sb.AppendLine($"ms.devlang: {MsDevLang}");
            if (ZonePivotGroups != null) sb.AppendLine($"zone_pivot_groups: {ZonePivotGroups}");

            if (sb.Length > 0)
            {
                var yamlDivider = new string('-', 3);
                sb.Insert(0, yamlDivider + Environment.NewLine);
                sb.AppendLine(yamlDivider);
            }
            return sb.ToString();
        }
    }
    
    internal static DocsMetadata Create(HtmlDocument htmlDoc)
    {
        var metadata = new DocsMetadata();

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
                case "persistent_id":
                    metadata.Id = value;
                    break;
                case "ms.assetid":
                    metadata.AssetId = value;
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
                case "page_type":
                    metadata.PageKind = value;
                    break;
                case "updated_at":
                    metadata.UpdatedAt = value;
                    break;
                default:
                    break;
            }
        }

        return metadata;
    }

    public string? Id { get; set; }
    public string? ZonePivotGroups { get; set; }
    public string? Author { get; set; }
    public string? MsAuthor { get; set; }
    public string? AssetId { get; set; }
    public string? Title { get; set; }
    public string? Description { get; set; }
    public string? Keywords { get; set; }
    public string? GitUrl { get; set; }
    public string? Manager { get; set; }
    public string? MsCustom { get; set; }
    public string? MsDate { get; set; }
    public string? MsService { get; set; }
    public string? MsDevLang { get; set; }
    public string? MsTopic { get; set; }
    public string? PageKind { get; set; }
    public string? Schema { get; set; }
    public string? UpdatedAt { get; set; }
}