using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docx.Renderer.Markdown.Renderers;
using DXPlus;

namespace Docx.Renderer.Markdown
{
    public class MarkdownRenderer : IMarkdownRenderer
    {
        private readonly List<IMarkdownObjectRenderer> renderers;

        public MarkdownRenderer()
        {
            renderers = new List<IMarkdownObjectRenderer>()
            {
                new ParagraphRenderer(),
                new RunRenderer()
            };
        }

        public string MediaFolder { get; private set; }

        public void Convert(string docxFile, string markdownFile, string mediaFolder)
        {
            if (string.IsNullOrEmpty(docxFile))
                throw new ArgumentException("Value cannot be null or empty.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"{docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrEmpty(markdownFile))
                throw new ArgumentException("Value cannot be null or empty.", nameof(markdownFile));
            if (string.IsNullOrEmpty(mediaFolder))
                throw new ArgumentException("Value cannot be null or empty.", nameof(mediaFolder));

            MediaFolder = mediaFolder;
            if (!Directory.Exists(mediaFolder))
                Directory.CreateDirectory(mediaFolder);

            if (File.Exists(markdownFile))
                File.Delete(markdownFile);

            using var document = Document.Load(docxFile);
            using var writer = new StreamWriter(markdownFile);

            var renderer = this as IMarkdownRenderer;
            renderer.WriteContainer(writer, document.Blocks, null);
        }

        public IMarkdownObjectRenderer FindRenderer(object element)
            => renderers.FirstOrDefault(r => r.CanRender(element));

        public void WriteContainer(TextWriter writer, IEnumerable<object> container, object tags)
        {
            foreach (var block in container)
            {
                WriteElement(writer, block, tags);
            }
        }

        public void WriteElement(TextWriter writer, object element, object tags)
        {
            var renderer = FindRenderer(element);
            if (renderer != null) renderer.Render(this, writer, element, tags);
            else Console.WriteLine($"Missing renderer for {element.GetType().Name}");
        }
    }
}
