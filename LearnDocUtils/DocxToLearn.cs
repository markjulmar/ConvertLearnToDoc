using System.Text.RegularExpressions;
using Docx.Renderer.Markdown;
using DXPlus;

namespace LearnDocUtils;

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

        string markdownFolder = Path.Combine(outputFolder, Constants.IncludesFolder);
        if (!Directory.Exists(markdownFolder))
            Directory.CreateDirectory(markdownFolder);

        options ??= new MarkdownOptions();

        string baseFilename = Path.GetFileNameWithoutExtension(docxFile);

        var markdownFile = Path.Combine(markdownFolder, baseFilename + "-temp.g.md");
        var mediaFolder = Path.Combine(outputFolder, Constants.MediaFolder);

        try
        {
            // Grab some pre-conversion options.
            using (var doc = Document.Load(docxFile))
            {
                if (options.Debug)
                {
                    await File.WriteAllTextAsync(
                        Path.Combine(outputFolder, baseFilename + "-temp.g.txt"),
                        DocxDebug.Dump(doc));

                    await File.WriteAllTextAsync(
                        Path.Combine(outputFolder, baseFilename + "-temp.g.xml"),
                        DocxDebug.FormatXml(doc));
                }

                if (doc.CustomProperties.TryGetValue(nameof(MarkdownOptions.UseAsterisksForBullets), out var yesNo))
                    options.UseAsterisksForBullets = yesNo?.Value == "True";
                if (doc.CustomProperties.TryGetValue(nameof(MarkdownOptions.UseAsterisksForEmphasis), out yesNo))
                    options.UseAsterisksForEmphasis = yesNo?.Value == "True";
            }

            // Convert the docx file to a single .md file
            new MarkdownRenderer(options).Convert(docxFile, markdownFile, mediaFolder);

            // Do some post-processing.
            var markdownText = DocToMarkdownRenderer.PostProcessMarkdown(await File.ReadAllTextAsync(markdownFile));
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
            if (!options.Debug)
            {
                if (File.Exists(markdownFile))
                    File.Delete(markdownFile);
            }
        }
    }
}