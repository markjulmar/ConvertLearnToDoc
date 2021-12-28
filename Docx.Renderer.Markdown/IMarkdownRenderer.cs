using System.Collections.Generic;
using System.IO;

namespace Docx.Renderer.Markdown
{
    public interface IMarkdownRenderer
    {
        IMarkdownObjectRenderer FindRenderer(object element);
        void WriteContainer(TextWriter writer, IEnumerable<object> container, object tags);
        void WriteElement(TextWriter writer, object element, object tags);
    }
}
