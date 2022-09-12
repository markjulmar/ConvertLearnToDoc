using System.Diagnostics;

namespace Markdig.Renderer.Docx.Inlines;

public class NolocInlineRenderer : DocxObjectRenderer<NolocInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, NolocInline nolocText)
    {
        Debug.Assert(currentParagraph != null);

        // Nothing to render .. ignore.
        if (string.IsNullOrEmpty(nolocText.Text))
            return;

        // Rendered (or will be) by another element?
        if (owner.OutOfPlaceRendered.Contains(nolocText))
        {
            owner.OutOfPlaceRendered.Remove(nolocText);
            return;
        }

        string text = Helpers.CleanText(nolocText.Text);

        //TODO: add comment indicating no localization should occur
        if (currentParagraph.Text.Length == 0)
        {
            currentParagraph.Text = text;
        }
        else
        {
            currentParagraph.AddText(text);
        }
    }
}