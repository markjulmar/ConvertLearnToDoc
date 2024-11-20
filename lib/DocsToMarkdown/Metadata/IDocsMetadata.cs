namespace Julmar.DocsToMarkdown.Metadata;

public interface IDocsMetadata
{
    string? PageKind { get; }
    string YamlHeader { get; }
    IReadOnlyDictionary<string,string> Values { get; }
}