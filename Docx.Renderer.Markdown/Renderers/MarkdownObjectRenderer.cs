using Julmar.GenMarkdown;

namespace Docx.Renderer.Markdown.Renderers
{
    public abstract class MarkdownObjectRenderer<T> : IMarkdownObjectRenderer
    {
        public bool CanRender(object element) => element is T;

        void IMarkdownObjectRenderer.Render(IMarkdownRenderer renderer,
            MarkdownDocument document, MarkdownBlock blockOwner,
            object elementToRender, RenderBag tags) =>
            Render(renderer, document, blockOwner, (T) elementToRender, tags);
        
        protected abstract void Render(IMarkdownRenderer renderer,
            MarkdownDocument document, MarkdownBlock blockOwner,
            T elementToRender, RenderBag tags);
    }
}
