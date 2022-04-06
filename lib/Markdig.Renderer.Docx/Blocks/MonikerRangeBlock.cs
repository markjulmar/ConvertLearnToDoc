using System.Drawing;

namespace Markdig.Renderer.Docx.Blocks;

public sealed class MonikerRangeRenderer : DocxObjectRenderer<MonikerRangeBlock>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MonikerRangeBlock block)
    {
        var p = currentParagraph ?? document.AddParagraph();
        p.AddText(new Run($"{{moniker: \"{block.MonikerRange}\"}}", new Formatting { Highlight = Highlight.DarkMagenta, Color = Color.White }));

        WriteChildren(block, owner, document, null);

        p = document.AddParagraph();
        p.AddText(new Run($"{{moniker-end: \"{block.MonikerRange}\"}}", new Formatting { Highlight = Highlight.DarkMagenta, Color = Color.White }));
        p.Newline();
    }
}