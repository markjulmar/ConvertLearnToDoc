using Julmar.GenMarkdown;

namespace LearnDocUtils;

public class MarkdownOptions : MarkdownFormatting
{
    public bool Debug { get; set; }
    public bool UsePlainMarkdown { get; set; }
    public string Metadata { get; set; }
    public bool IgnoreEmbeddedMetadata { get; set; }
    public bool UseGenericIds { get; set; }
}
