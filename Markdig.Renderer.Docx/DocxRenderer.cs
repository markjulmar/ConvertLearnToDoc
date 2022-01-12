using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DXPlus;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Renderer.Docx.Inlines;
using Markdig.Syntax;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using TripleColonInlineRenderer = Markdig.Renderer.Docx.Inlines.TripleColonInlineRenderer;

namespace Markdig.Renderer.Docx
{
    /// <summary>
    /// DoxC renderer for a Markdown <see cref="MarkdownDocument"/> object.
    /// </summary>
    public class DocxObjectRenderer : IDocxRenderer
    {
        private readonly IDocument document;
        private readonly List<IDocxObjectRenderer> renderers;
        private readonly string moduleFolder;
        private readonly Action<string> logger;

        /// <summary>
        /// This holds elements where previous inline renderers had to reach into the stream
        /// and render siblings. It's used to avoid double rendering.
        /// </summary>
        public IList<MarkdownObject> OutOfPlaceRendered => new List<MarkdownObject>();

        public string ZonePivot { get; }

        public DocxObjectRenderer(IDocument document, string moduleFolder, string zonePivot, Action<string> logger = null)
        {
            this.moduleFolder = moduleFolder;
            this.document = document;
            this.ZonePivot = zonePivot;
            this.logger = logger;

            renderers = new List<IDocxObjectRenderer>
            {
                // Block handlers
                new HeadingRenderer(),
                new ParagraphRenderer(),
                new ListRenderer(),
                new QuoteBlockRenderer(),
                new QuoteSectionNoteRenderer(),
                new FencedCodeBlockRenderer(),
                new TripleColonRenderer(),
                new FencedCodeBlockRenderer(),
                new TableRenderer(),
                new InclusionRenderer(),

                // Inline handlers
                new LiteralInlineRenderer(),
                new EmphasisInlineRenderer(),
                new LineBreakInlineRenderer(),
                new LinkInlineRenderer(),
                new AutolinkInlineRenderer(),
                new CodeInlineRenderer(),
                new DelimiterInlineRenderer(),
                new HtmlEntityInlineRenderer(),
                new LinkReferenceDefinitionRenderer(),
                new TaskListRenderer(),
                new HtmlInlineRenderer(),
                new TripleColonInlineRenderer()
            };
        }

        public IDocxObjectRenderer FindRenderer(MarkdownObject obj)
        {
            var renderer = renderers.FirstOrDefault(r => r.CanRender(obj));
            if (renderer == null)
                logger?.Invoke($"Missing renderer for {obj.GetType()}");
            return renderer;
        }

        public void Render(MarkdownDocument markdownDocument)
        {
            for (var index = 0; index < markdownDocument.Count; index++)
            {
                var block = markdownDocument[index];

                // Special case RowBlock and children to generate a full table.
                if (block is RowBlock)
                {
                    var rows = new List<RowBlock>();
                    do
                    {
                        rows.Add((RowBlock) block);
                        block = markdownDocument[++index];
                    } while (block is RowBlock);

                    new RowBlockRenderer().Write(this, document, rows);
                }

                // Find the renderer and process.
                var renderer = FindRenderer(block);

                try
                {
                    renderer?.Write(this, document, null, block);
                }
                catch (AggregateException aex)
                {
                    var ex = aex.Flatten();
                    logger?.Invoke($"{ex.GetType().Name}: {ex.Message}");
                    logger?.Invoke(ex.StackTrace);
                }
                catch (Exception ex)
                {
                    logger?.Invoke($"{ex.GetType().Name}: {ex.Message}");
                    logger?.Invoke(ex.StackTrace);
                }
            }
        }

        public Drawing InsertImage(Paragraph currentParagraph, string imageSource, string altText)
        {
            string path = ResolvePath(moduleFolder, imageSource);
            if (File.Exists(path))
            {
                var image = document.AddImage(path);
                var drawing = image.CreatePicture(imageSource, altText);
                currentParagraph.Append(drawing);
                return drawing;
            }

            return null;
        }

        /// <summary>
        /// Returns a specific embedded resource by name.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public Stream GetEmbeddedResource(string name) => Assembly.GetExecutingAssembly().GetManifestResourceStream("Markdig.Renderer.Docx.Resources."+name);

        private static string ResolvePath(string rootFolder, string path)
            => path.Contains(':')
                ? path
                : Path.IsPathRooted(path)
                    ? path
                    : Path.Combine(rootFolder, path);

    }
}