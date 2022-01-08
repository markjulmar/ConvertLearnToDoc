using System;
using System.IO;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class DocxToLearn
    {
        public static async Task ConvertAsync(string docxFile, string outputFolder,
            bool debug, IDocxToMarkdown converter)
        {
            if (string.IsNullOrWhiteSpace(docxFile))
                throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or whitespace.", nameof(outputFolder));
            if (converter is null)
                throw new ArgumentNullException(nameof(converter));
 
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            string markdownFile = Path.Combine(outputFolder, "temp.md");
            string mediaFolder = Path.Combine(outputFolder, "media");

            try
            {
                // Convert the docx file to a single .md file
                await converter.ConvertAsync(docxFile, markdownFile, mediaFolder);

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
                if (!debug)
                {
                    File.Delete(markdownFile);
                }
            }
        }
    }
}