namespace Docx.Renderer.Markdown;

public class DocxMarkdownFormatting : MarkdownFormatting
{
    public bool PreferPlainMarkdown { get; set; }
    public Func<string,string> ConvertAbsoluteUrls { get; set; }
}