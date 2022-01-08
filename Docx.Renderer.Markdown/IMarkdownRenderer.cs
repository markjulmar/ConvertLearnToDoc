using System.Collections.Generic;
using Julmar.GenMarkdown;

namespace Docx.Renderer.Markdown
{
    public interface IMarkdownRenderer
    {
        string MediaFolder { get; }
        IMarkdownObjectRenderer FindRenderer(object element);
        void WriteContainer(MarkdownDocument document, MarkdownBlock blockOwner, IEnumerable<object> container, RenderBag tags);
        void WriteElement(MarkdownDocument document, MarkdownBlock blockOwner, object element, RenderBag tags);
    }
}
