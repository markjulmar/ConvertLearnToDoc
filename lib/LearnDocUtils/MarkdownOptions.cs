using Julmar.GenMarkdown;

namespace LearnDocUtils;

public class MarkdownOptions : MarkdownFormatting
{
    public bool Debug { get; set; }
    public bool UsePlainMarkdown { get; set; }
}

public class LearnMarkdownOptions : MarkdownOptions
{
    public bool IgnoreMetadata { get; set; }
    public bool UseGenericIds { get; set; }
}