namespace Markdig.Renderer.Docx;

public interface IDocxRenderer
{
    IList<MarkdownObject> OutOfPlaceRendered { get; }
    string ZonePivot { get; }
    IDocxObjectRenderer FindRenderer(MarkdownObject obj);
    Drawing InsertImage(Paragraph currentParagraph, MarkdownObject source, string imageUrl, string altText, string title, bool hasBorder);
    Stream GetEmbeddedResource(string name);
    void AddComment(Paragraph owner, string commentText);
    byte[] GetFile(MarkdownObject source, string path);
}