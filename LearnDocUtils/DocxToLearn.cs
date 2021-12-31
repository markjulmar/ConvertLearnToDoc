using System;
using System.IO;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class DocxToLearn
    {
        public static async Task ConvertAsync(string docxFile, string outputFolder,
            Action<string> logger, bool debug, bool usePandoc)
        {
            if (string.IsNullOrWhiteSpace(docxFile))
                throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or whitespace.", nameof(outputFolder));

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            IDocxToMarkdown converter = usePandoc ? new DocxToMarkdownPandoc() : new DocxToMarkdownDxPlus();
            string markdownFile = Path.Combine(outputFolder, "temp.md");
            string mediaFolder = Path.Combine(outputFolder, "media");

            try
            {
                // Convert the docx file to a single .md file
                await converter.ConvertAsync(docxFile, markdownFile, mediaFolder, logger, debug);

                // Now build a module from the markdown contents
                var moduleBuilder = new ModuleBuilder(docxFile, outputFolder, markdownFile, logger);
                await moduleBuilder.CreateModuleAsync(converter.MarkdownProcessor);
            }
            catch
            {
                Directory.Delete(outputFolder, true);
                throw;
            }
            finally
            {
                if (!debug)
                {
                    File.Delete(markdownFile);
                }
            }
        }
    }
}