namespace Docx.Renderer.Markdown.Renderers;

public class IgnoredBlocks : IMarkdownObjectRenderer
{
    public bool CanRender(object element)
    {
        // Ignore comments/bookmarks
        return element is UnknownBlock ub && 
               (ub.Name.Contains("comment") || ub.Name.Contains("bookmark"));
    }

    public void Render(IMarkdownRenderer renderer, MarkdownDocument document, MarkdownBlock blockOwner, object elementToRender, RenderBag tags)
    {
    }
}