namespace Docx.Renderer.Markdown;

public class DocxMarkdownFormatting : MarkdownFormatting
{
    public Func<string,string> ConvertAbsoluteUrls { get; set; }
}