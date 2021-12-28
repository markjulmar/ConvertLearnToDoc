using System.IO;

namespace Docx.Renderer.Markdown
{
    public interface IMarkdownObjectRenderer
    {
        bool CanRender(object element);
        void Render(IMarkdownRenderer renderer, TextWriter writer, object element, object tags);
    }
}
