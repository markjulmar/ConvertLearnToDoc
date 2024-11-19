namespace Julmar.DocsToMarkdown.Metadata;

public interface IDocsMetadata
{
    string? Id { get; set; }
    string? ZonePivotGroups { get; set; }
    string? Author { get; set; }
    string? MsAuthor { get; set; }
    string? Title { get; set; }
    string? Description { get; set; }
    string? Keywords { get; set; }
    string? GitUrl { get; set; }
    string? Manager { get; set; }
    string? MsCustom { get; set; }
    string? MsDate { get; set; }
    string? MsService { get; set; }
    string? MsDevLang { get; set; }
    string? MsTopic { get; set; }
    string? PageKind { get; set; }
    string? UpdatedAt { get; set; }

    string YamlHeader { get; }
}