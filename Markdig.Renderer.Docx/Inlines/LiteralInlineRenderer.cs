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

        if (currentParagraph.Text.Length == 0)
        {
            currentParagraph.SetText(text);
        }
        else
        {
            currentParagraph.Append(text);
        }
    }
}