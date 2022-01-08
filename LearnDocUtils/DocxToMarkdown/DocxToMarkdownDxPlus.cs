using System.Threading.Tasks;
using Docx.Renderer.Markdown;

namespace LearnDocUtils
{
    sealed class DocxToMarkdownDxPlus : IDocxToMarkdown
    {
        public Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder)
        {
            new MarkdownRenderer().Convert(docxFile, markdownFile, mediaFolder);
            return Task.CompletedTask;
        }
    }
}
