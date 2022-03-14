using Markdig.Helpers;

namespace Markdig.Renderer.Docx;

public abstract class DocxObjectRenderer<TObject> : IDocxObjectRenderer
    where TObject : MarkdownObject 
{
    public virtual bool CanRender(MarkdownObject obj) => obj.GetType() == typeof(TObject) || obj is TObject;
    public abstract void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TObject obj);
    void IDocxObjectRenderer.Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, MarkdownObject obj)
        => Write(owner, document, currentParagraph, (TObject)obj);

    public void WriteChildren(LeafBlock leafBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
    {
        var inlines = leafBlock.Inline;
        if (inlines != null)
        {
            foreach (var child in inlines)
            {
                Write(child, owner, document, currentParagraph);
            }
        }

        if (leafBlock.Lines.Count > 0)
        {
            int index = 0;
            int count = leafBlock.Lines.Count;
            foreach (var text in leafBlock.Lines.Cast<StringLine>().Take(count))
            {
                currentParagraph.AddRange(Run.Create(Helpers.CleanText(text.ToString())));
                if (++index < count)
                    currentParagraph.Newline();
            }
        }
    }

    public void WriteChildren(ContainerBlock container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
    {
        foreach (var block in container)
        {
            Write(block, owner, document, currentParagraph);
        }
    }

    public void WriteChildren(ContainerInline container, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
    {
        foreach (var inline in container)
        {
            Write(inline, owner, document, currentParagraph);
        }
    }

    public void Write(MarkdownObject item, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
    {
        var renderer = owner.FindRenderer(item);
        renderer?.Write(owner, document, currentParagraph, item);
    }
}