using DXPlus;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx
{
    public interface IDocxObjectRenderer
    {
        bool CanRender(MarkdownObject obj);
        void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MarkdownObject obj);
        void WriteChildren(LeafBlock leafBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void WriteChildren(ContainerBlock container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void WriteChildren(ContainerInline container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
        void Write(MarkdownObject item, IDocxRenderer owner, IDocument document, Paragraph currentParagraph);
    }
}