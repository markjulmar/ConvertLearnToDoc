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
        private string mediaFolder;
        private List<IMarkdownObjectRenderer> renderers;

        public MarkdownRenderer()
        {
            renderers = new List<IMarkdownObjectRenderer>()
            {
                new ParagraphRenderer(),
                new RunRenderer()
            };
        }

        public string MediaFolder => mediaFolder;

        public void Convert(string docxFile, string markdownFile, string mediaFolder)
        {
            using var document = Document.Load(docxFile);
            using var writer = new StreamWriter(markdownFile);

            if (string.IsNullOrEmpty(mediaFolder))
                mediaFolder = Path.GetDirectoryName(markdownFile);
            else if (!Path.IsPathRooted(mediaFolder))
                mediaFolder = Path.Combine(Path.GetDirectoryName(markdownFile), mediaFolder);
            if (!Directory.Exists(mediaFolder))
                Directory.CreateDirectory(mediaFolder);

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
