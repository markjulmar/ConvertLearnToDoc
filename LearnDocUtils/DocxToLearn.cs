using System;
using System.IO;
using System.Threading.Tasks;
using Docx.Renderer.Markdown;

namespace LearnDocUtils
{
    public static class DocxToLearn
    {
        public static async Task ConvertAsync(string docxFile, string outputFolder, MarkdownOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(docxFile))
                throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or whitespace.", nameof(outputFolder));
 
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var markdownFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxFile)+"-temp.g.md");
            var mediaFolder = Path.Combine(outputFolder, "media");

            try
            {
                // Convert the docx file to a single .md file
                new MarkdownRenderer(options).Convert(docxFile, markdownFile, mediaFolder);

                // Do some post-processing.
                var markdownText = PostProcessMarkdown(await File.ReadAllTextAsync(markdownFile));
                await File.WriteAllTextAsync(markdownFile, markdownText);

                // Now build a module from the markdown contents
                var moduleBuilder = new ModuleBuilder(docxFile, outputFolder, markdownFile);
                await moduleBuilder.CreateModuleAsync();
            }
            catch
            {
                Directory.Delete(outputFolder, true);
                throw;
            }
            finally
            {
                if (options == null || options.Debug == false)
                {
                    if (File.Exists(markdownFile))
                        File.Delete(markdownFile);
                }
            }
        }

        private static string PostProcessMarkdown(string text)
        {
            text = text.Trim('\r').Trim('\n');
            text = text.Replace("(media/", "(../media/");
            text = text.Replace(LearnUtilities.AbsolutePathMarker, string.Empty);

            return text;
        }

    }
}