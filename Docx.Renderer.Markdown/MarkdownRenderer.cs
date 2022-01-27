using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docx.Renderer.Markdown.Renderers;
using DXPlus;
using Julmar.GenMarkdown;

namespace Docx.Renderer.Markdown
{
    public class MarkdownRenderer : IMarkdownRenderer
    {
        private readonly MarkdownFormatting options;
        private readonly List<IMarkdownObjectRenderer> renderers;

        public MarkdownRenderer() : this (null)
        {
        }

        public MarkdownRenderer(MarkdownFormatting options)
        {
            this.options = options;
            renderers = new List<IMarkdownObjectRenderer>
            {
                new ParagraphRenderer(),
                new TableRenderer(),
                new RunRenderer()
            };
        }

        public string RelativeMediaFolder { get; private set; }
        public string MediaFolder { get; private set; }

        public void Convert(string docxFile, string markdownFile, string mediaFolder)
        {
            if (string.IsNullOrEmpty(docxFile))
                throw new ArgumentException("Value cannot be null or empty.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"{docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrEmpty(markdownFile))
                throw new ArgumentException("Value cannot be null or empty.", nameof(markdownFile));
            if (Path.GetInvalidPathChars().Any(markdownFile.Contains))
                throw new ArgumentException($"{markdownFile} is an invalid filename.", nameof(markdownFile));
            if (string.IsNullOrEmpty(mediaFolder))
                throw new ArgumentException("Value cannot be null or empty.", nameof(mediaFolder));

            MediaFolder = mediaFolder;
            RelativeMediaFolder = Path.GetRelativePath(Path.GetDirectoryName(markdownFile)??"", mediaFolder);

            if (!Directory.Exists(MediaFolder))
                Directory.CreateDirectory(MediaFolder);

            if (File.Exists(markdownFile))
                File.Delete(markdownFile);

            using var wordDocument = Document.Load(docxFile);
            var markdownDocument = new MarkdownDocument();

            var renderer = this as IMarkdownRenderer;
            renderer.WriteContainer(markdownDocument, null, wordDocument.Blocks, null);

            using var writer = new StreamWriter(markdownFile);
            markdownDocument.Write(writer, options);
        }

        public IMarkdownObjectRenderer FindRenderer(object element)
            => renderers.FirstOrDefault(r => r.CanRender(element));

        public void WriteContainer(MarkdownDocument document, MarkdownBlock blockOwner, IEnumerable<object> container, RenderBag tags)
        {
            foreach (var block in container)
            {
                WriteElement(document, blockOwner, block, tags);
            }
        }

        public void WriteElement(MarkdownDocument document, MarkdownBlock blockOwner, 
                                    object element, RenderBag tags)
        {
            var renderer = FindRenderer(element);
            if (renderer != null) 
                renderer.Render(this, document, blockOwner, element, tags);
            else 
                Console.WriteLine($"Missing renderer for {element.GetType().Name}");
        }
    }
}
