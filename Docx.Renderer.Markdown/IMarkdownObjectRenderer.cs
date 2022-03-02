namespace Docx.Renderer.Markdown;

public interface IMarkdownObjectRenderer
{
    bool CanRender(object element);
    void Render(IMarkdownRenderer renderer, 
        MarkdownDocument document, MarkdownBlock blockOwner,
        object elementToRender, RenderBag tags);
}