namespace Markdig.Renderer.Docx.Inlines;

public class AutolinkInlineRenderer : DocxObjectRenderer<AutolinkInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, AutolinkInline link)
    {
        string url = link.Url;
        currentParagraph.Append(new Hyperlink(url, new Uri(url, UriKind.Absolute)));
    }
}