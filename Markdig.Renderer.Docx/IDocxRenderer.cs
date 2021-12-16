using System.Collections.Generic;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx
{
    public interface IDocxRenderer
    {
        IList<MarkdownObject> OutOfPlaceRendered { get; }
        string ZonePivot { get; }
        IDocxObjectRenderer FindRenderer(MarkdownObject obj);
        Picture InsertImage(Paragraph currentParagraph, string imageSource, string altText);
    }
}