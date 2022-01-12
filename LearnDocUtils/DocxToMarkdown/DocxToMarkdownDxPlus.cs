using System.IO;
using System.Threading.Tasks;
using Docx.Renderer.Markdown;

namespace LearnDocUtils
{
    sealed class DocxToMarkdownDxPlus : IDocxToMarkdown
    {
        public async Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder)
        {
            new MarkdownRenderer().Convert(docxFile, markdownFile, mediaFolder);
            
            // Do some post-processing.
            var markdownText = PostProcessMarkdown(await File.ReadAllTextAsync(markdownFile));
            await File.WriteAllTextAsync(markdownFile, markdownText);
        }

        private static string PostProcessMarkdown(string text)
        {
            text = text.Trim('\r').Trim('\n');
            text = text.Replace("(media/", "(../media/");

            return text;
        }
    }
}
