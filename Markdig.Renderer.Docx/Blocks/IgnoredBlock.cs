namespace Markdig.Renderer.Docx.Blocks;

public sealed class IgnoredBlock : IDocxObjectRenderer
{
    private readonly Type blockType;
    public IgnoredBlock(Type blockType) { this.blockType = blockType; }
    public bool CanRender(MarkdownObject obj) => obj?.GetType().IsAssignableTo(blockType) == true;
    public void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MarkdownObject obj) {}
    public void WriteChildren(LeafBlock leafBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph) { }
    public void WriteChildren(ContainerBlock container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph) { }
    public void WriteChildren(ContainerInline container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph) { }
    public void Write(MarkdownObject item, IDocxRenderer owner, IDocument document, Paragraph currentParagraph) { }
}