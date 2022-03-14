namespace Markdig.Renderer.Docx.Inlines;

public class DelimiterInlineRenderer : DocxObjectRenderer<DelimiterInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, DelimiterInline obj)
    {
        currentParagraph.Add(obj.ToLiteral());
    }
}