using Markdig.Helpers;
using Markdig.Parsers;
using Markdig.Renderers;

namespace Markdig.Renderer.Docx.TripleColonExtensions;

internal sealed class DocxTripleColonExtension(MarkdownContext context) : IMarkdownExtension
{
    public void Setup(MarkdownPipelineBuilder pipeline)
    {
        new TripleColonExtension(context).Setup(pipeline);
        var parser = pipeline.InlineParsers.Find<TripleColonInlineParser>();
        pipeline.InlineParsers.Replace<TripleColonInlineParser>(new InlineImageParser(parser));
    }

    public void Setup(MarkdownPipeline pipeline, IMarkdownRenderer renderer)
    {
        new TripleColonExtension(context).Setup(pipeline, renderer);
    }

    private sealed class InlineImageParser : InlineParser
    {
        private readonly TripleColonInlineParser parser;

        public InlineImageParser(TripleColonInlineParser parser)
        {
            this.parser = parser;
            OpeningCharacters = parser.OpeningCharacters;
        }

        public override bool Match(InlineProcessor processor, ref StringSlice slice)
        {
            if (!parser.Match(processor, ref slice))
                return false;

            // DocFX v2 leaves inline images open as containers. Close them before
            // Markdig parses following icons or text, so those remain siblings.
            if (processor.Inline is TripleColonInline { Extension.Name: "image" } image)
                image.IsClosed = true;

            return true;
        }
    }
}
