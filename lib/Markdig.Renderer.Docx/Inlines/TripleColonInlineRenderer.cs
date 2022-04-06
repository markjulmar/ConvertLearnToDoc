using Markdig.Renderer.Docx.TripleColonExtensions;

namespace Markdig.Renderer.Docx.Inlines;

public class TripleColonInlineRenderer : DocxObjectRenderer<TripleColonInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonInline inline)
    {
        TripleColonProcessor.Write(this, inline, owner, document, currentParagraph, new TripleColonElement(inline));
    }
}