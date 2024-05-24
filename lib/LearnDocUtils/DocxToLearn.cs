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

        var conversionOptions = new DocxMarkdownFormatting
        {
            OrderedListUsesSequence = options?.OrderedListUsesSequence ?? false,
            UseAlternateHeaderSyntax = false, // never allowed
            UseAsterisksForEmphasis = options?.UseAsterisksForEmphasis ?? false,
            UseAsterisksForBullets = options?.UseAsterisksForBullets ?? false,
            UseIndentsForCodeBlocks = options?.UseIndentsForCodeBlocks ?? false,
            EscapeAllIntrawordEmphasis = options?.EscapeAllIntrawordEmphasis ?? false,
            PrettyPipeTables = options?.PrettyPipeTables ?? false,
            PreferPlainMarkdown = options?.UsePlainMarkdown ?? false,
        };

        string baseFilename = Path.GetFileNameWithoutExtension(docxFile);
        var markdownFile = Path.Combine(markdownFolder, baseFilename + "-temp.g.md");
        var mediaFolder = Path.Combine(outputFolder, Constants.MediaFolder);
        string baseUrl = "";

        try
        {
            // Grab some pre-conversion options.
            using (var doc = Document.Load(docxFile))
            {
                if (options?.Debug == true)
                {
                    await File.WriteAllTextAsync(
                        Path.Combine(outputFolder, baseFilename + "-temp.g.txt"),
                        DocxDebug.Dump(doc));

                    await File.WriteAllTextAsync(
                        Path.Combine(outputFolder, baseFilename + "-temp.g.xml"),
                        DocxDebug.FormatXml(doc));
                }

                if (doc.CustomProperties.TryGetValue(nameof(MarkdownOptions.UseAsterisksForBullets), out var yesNo))
                    conversionOptions.UseAsterisksForBullets = yesNo?.Value == "True";
                if (doc.CustomProperties.TryGetValue(nameof(MarkdownOptions.UseAsterisksForEmphasis), out yesNo))
                    conversionOptions.UseAsterisksForEmphasis = yesNo?.Value == "True";
                baseUrl = doc.Properties.Description;
            }

            // Logic to convert urls.
            conversionOptions.ConvertAbsoluteUrls = url => ConvertUrls.FromAbsolute(url, baseUrl);

            // Convert the docx file to a single .md file
            new MarkdownRenderer(conversionOptions).Convert(docxFile, markdownFile, mediaFolder);

            // Do some post-processing.
            var markdownText = DocToMarkdownRenderer.PostProcessMarkdown(await File.ReadAllTextAsync(markdownFile));
            await File.WriteAllTextAsync(markdownFile, markdownText);

            // Now build a module from the markdown contents
            var moduleBuilder = new ModuleBuilder(docxFile, outputFolder, markdownFile);
            await moduleBuilder.CreateModuleAsync(new LearnModuleOptions
            {
                IgnoreEmbeddedMetadata  = options?.IgnoreEmbeddedMetadata ?? false, 
                Metadata = options?.Metadata,
                UseGenericIds = options?.UseGenericIds ?? false
            });
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
}