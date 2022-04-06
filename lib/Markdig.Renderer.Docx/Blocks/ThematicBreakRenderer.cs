using System.Drawing;

namespace Markdig.Renderer.Docx.Blocks;

/// <summary>
/// This renders a thematic break - these are used in tab conceptual content to
/// indicate the end of the group as a {---} marker.
/// </summary>
public sealed class ThematicBreakRenderer : DocxObjectRenderer<ThematicBreakBlock>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, ThematicBreakBlock block)
    {
        if (block.ThematicChar == '-' && block.ThematicCharCount == 3)
        {
            currentParagraph ??= document.AddParagraph();
            currentParagraph.AddText(new Run("{tabgroup-end}", 
                new Formatting { Highlight = Highlight.DarkGray, Color = Color.White }));
        }
    }
}