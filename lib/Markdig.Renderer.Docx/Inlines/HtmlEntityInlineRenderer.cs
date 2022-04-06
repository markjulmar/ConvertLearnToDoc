namespace Markdig.Renderer.Docx.Inlines;

public class HtmlEntityInlineRenderer : DocxObjectRenderer<HtmlEntityInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlEntityInline obj)
    {
        var slice = obj.Transcoded;
        currentParagraph.AddText(slice.Text.Substring(slice.Start, slice.Length));
    }
}