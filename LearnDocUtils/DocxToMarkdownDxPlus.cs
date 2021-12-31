using System;
using System.Threading.Tasks;
using Docx.Renderer.Markdown;

namespace LearnDocUtils
{
    public class DocxToMarkdownDxPlus : IDocxToMarkdown
    {
        public Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder, Action<string> logger, bool debug)
        {
            new MarkdownRenderer().Convert(docxFile, markdownFile, mediaFolder);
            return Task.CompletedTask;
        }

        public Func<string, string> MarkdownProcessor => null;
    }
}
