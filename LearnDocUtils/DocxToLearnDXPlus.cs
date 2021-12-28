using System;
using System.IO;
using System.Threading.Tasks;
using Docx.Renderer.Markdown;

namespace LearnDocUtils
{
    public class DocxToLearnDXPlus : IDocxToLearn
    {
        public Task ConvertAsync(string docxFile, string outputFolder, Action<string> logger = null, bool debug = false)
        {
            string markdownFile = Path.Combine(outputFolder, "temp.md");
            string mediaFolder = Path.Combine(outputFolder, "media");

            try
            {
                var markdownConverter = new MarkdownRenderer();
                markdownConverter.Convert(docxFile, markdownFile, mediaFolder);
            }
            finally
            {
                if (debug)
                {
                    File.Delete(markdownFile);
                }
            }

            return Task.CompletedTask;
        }
    }
}
