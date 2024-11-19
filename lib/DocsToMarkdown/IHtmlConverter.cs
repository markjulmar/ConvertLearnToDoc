using Julmar.DocsToMarkdown.Metadata;

namespace Julmar.DocsToMarkdown;

public interface IHtmlConverter
{
    public Action<string>? UnhandledTagCallback { get; set; }
    public bool UseDocfxExtensions { get; set; }
    public IDocsMetadata Metadata { get; }
    IEnumerable<string> Process();
}