using Markdig.Renderer.Docx.TripleColonExtensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class TripleColonRenderer : DocxObjectRenderer<TripleColonBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonBlock block)
        {
            TripleColonProcessor.Write(this, block, owner, document, currentParagraph, new TripleColonElement(block));
        }
    }
}
