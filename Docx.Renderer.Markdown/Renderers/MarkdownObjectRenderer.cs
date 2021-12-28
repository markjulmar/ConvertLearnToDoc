using System.IO;

namespace Docx.Renderer.Markdown.Renderers
{
    public abstract class MarkdownObjectRenderer<T> : IMarkdownObjectRenderer
    {
        public bool CanRender(object element) => element is T;
        protected abstract void Render(IMarkdownRenderer renderer, TextWriter writer, T element, object tags);
        void IMarkdownObjectRenderer.Render(IMarkdownRenderer renderer, TextWriter writer, object element, object tags) => Render(renderer, writer, (T)element, tags);
    }
}
