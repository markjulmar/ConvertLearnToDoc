namespace Markdig.Renderer.Docx.Blocks;

public class HtmlBlockRenderer : DocxObjectRenderer<HtmlBlock>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlBlock block)
    {
        if (block.Type == HtmlBlockType.Comment)
        {
            // TODO: should we put this in as a comment to the doc?
        }
    }
}