using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Metadata;

public sealed class DocsMetadata : BaseMetadata
{
    internal DocsMetadata(HtmlDocument document) : base(document)
    {
    }

    public override string? PageKind => LoadedValues.GetValueOrDefault("page_type");
}