using DXPlus;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LineBreakInlineRenderer : DocxObjectRenderer<LineBreakInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LineBreakInline inline)
        {
            if (inline.IsHard)
                currentParagraph?.AppendLine();
            // Soft break - author didn't use word-wrap in the editor and put CRLF
            // in between lines. Add a space if the paragraph doesn't have one.
            else
            {
                if (currentParagraph != null && !currentParagraph.Text.EndsWith(' '))
                    currentParagraph.Append(" ");
            }
        }
    }
}