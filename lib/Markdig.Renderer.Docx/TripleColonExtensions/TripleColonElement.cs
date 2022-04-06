namespace Markdig.Renderer.Docx.TripleColonExtensions;

public sealed class TripleColonElement
{
    public ITripleColonExtensionInfo Extension { get; set; }
    public IDictionary<string, string> RenderProperties { get; set; }
    public bool Closed { get; set; }
    public bool EndingTripleColons { get; set; }
    public IDictionary<string, string> Attributes { get; set; }
    public ContainerBlock Container { get; set; }
    public ContainerInline Inlines { get; set; }

    public TripleColonElement(TripleColonBlock block)
    {
        Container = block;
        Extension = block.Extension;
        RenderProperties = block.RenderProperties;
        Closed = block.Closed;
        EndingTripleColons = block.EndingTripleColons;
        Attributes = block.Attributes;
    }

    public TripleColonElement(TripleColonInline inline)
    {
        Inlines = inline;
        Extension = inline.Extension;
        RenderProperties = inline.RenderProperties;
        Closed = inline.Closed;
        EndingTripleColons = inline.EndingTripleColons;
        Attributes = inline.Attributes;
    }
}