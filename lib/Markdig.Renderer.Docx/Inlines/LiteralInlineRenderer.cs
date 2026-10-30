using System.Diagnostics;

namespace Markdig.Renderer.Docx.Inlines;

public class LiteralInlineRenderer : DocxObjectRenderer<LiteralInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LiteralInline literal)
    {
        Debug.Assert(currentParagraph != null);

        // Nothing to render .. ignore.
        if (literal.Content.IsEmpty) 
            return;

        // Rendered (or will be) by another element?
        if (owner.OutOfPlaceRendered.Contains(literal))
        {
            owner.OutOfPlaceRendered.Remove(literal);
            return;
        }

        string text = Helpers.CleanText(literal.Content.ToString());

        // An image-only paragraph has no text, but replacing Text would remove
        // its drawing runs. Append so inline icons survive the following label.
        currentParagraph.AddText(text);
    }
}
