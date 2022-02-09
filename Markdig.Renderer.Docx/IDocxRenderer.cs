using System.Collections.Generic;
using System.IO;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx
{
    public interface IDocxRenderer
    {
        IList<MarkdownObject> OutOfPlaceRendered { get; }
        string ZonePivot { get; }
        IDocxObjectRenderer FindRenderer(MarkdownObject obj);
        Drawing InsertImage(Paragraph currentParagraph, string imageSource, string altText, string type, string localization, bool hasBorder, bool isLightbox);
        Stream GetEmbeddedResource(string name);
        void AddComment(Paragraph owner, string commentText);
    }
}