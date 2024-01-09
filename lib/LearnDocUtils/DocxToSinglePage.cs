using Docx.Renderer.Markdown;
using MSLearnRepos;
using ConvertLearnToDoc.Shared;
using Document = DXPlus.Document;

namespace LearnDocUtils;

public static class DocxToSinglePage
{
    public static async Task ConvertAsync(string docxFile, string markdownFile, MarkdownOptions options, bool preferPlainMarkdown)
    {
        if (string.IsNullOrWhiteSpace(docxFile))
            throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
        if (!File.Exists(docxFile))
            throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
        if (string.IsNullOrWhiteSpace(markdownFile))
            throw new ArgumentException($"'{nameof(markdownFile)}' cannot be null or whitespace.", nameof(markdownFile));

        var conversionOptions = new DocxMarkdownFormatting
        {
            OrderedListUsesSequence = options?.OrderedListUsesSequence ?? false,
            UseAlternateHeaderSyntax = false, // never allowed
            UseAsterisksForEmphasis = options?.UseAsterisksForEmphasis ?? false,
            UseAsterisksForBullets = options?.UseAsterisksForBullets ?? false,
            UseIndentsForCodeBlocks = options?.UseIndentsForCodeBlocks ?? false,
            EscapeAllIntrawordEmphasis = options?.EscapeAllIntrawordEmphasis ?? false,
            PrettyPipeTables = options?.PrettyPipeTables ?? false,
            PreferPlainMarkdown = preferPlainMarkdown
        };

        string baseFilename = Path.GetFileNameWithoutExtension(docxFile);
        string outputFolder = Path.GetDirectoryName(markdownFile) ?? Directory.GetCurrentDirectory();
        string mediaFolder = Path.Combine(outputFolder, Constants.MediaFolder);
        string baseUrl = null;

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
            if (doc.CustomProperties.TryGetValue(nameof(Uri), out var baseUri))
                baseUrl = baseUri?.Value;
        }

        conversionOptions.ConvertAbsoluteUrls = url =>
            !string.IsNullOrEmpty(baseUrl) && url.StartsWith(baseUrl) ? url[baseUrl.Length..] + ".md" : url;

        // Convert the docx file to a single .md file
        new MarkdownRenderer(conversionOptions).Convert(docxFile, markdownFile, mediaFolder);

        // Do some common post-processing.
        string markdownText = await File.ReadAllTextAsync(markdownFile);
        if (!preferPlainMarkdown)
        {
            markdownText = DocToMarkdownRenderer.PostProcessMarkdown(markdownText);
        }

        await File.WriteAllTextAsync(markdownFile, markdownText);

        Dictionary<object, object> suppliedMetadata = null;
        if (options?.Metadata != null)
        {
            try
            {
                suppliedMetadata = PersistenceUtilities.YamlStringToDictionary(options.Metadata);
            }
            catch
            {
                // Ignore
            }
        }

        await WriteMetadata(docxFile, markdownFile, suppliedMetadata);
    }

    private static async Task WriteMetadata(string docxFile, string markdownFile, Dictionary<object,object> suppliedMetadata)
    {
        string[] validHeaders = {"Title", "Subtitle", "Author", "Abstract"};
        var doc = Document.Load(docxFile);

        // Load any original metadata from the document.
        if (suppliedMetadata == null && 
            doc.CustomProperties.TryGetValue(nameof(Metadata), out var jsonText))
        {
            if (jsonText?.Value != null)
            {
                // Get the dictionary of values.
                try
                {
                    suppliedMetadata = PersistenceUtilities
                        .JsonStringToDictionary(jsonText.Value);
                }
                catch
                {
                    // Ignore
                }

            }
        }

        bool foundStart = false;
        string title = null, author = null, summary = null;

        foreach (var item in doc.Paragraphs)
        {
            var styleName = item.Properties.StyleName;
            if (styleName == "Heading1")
            {
                foundStart = true;
                break;
            }
            
            if (!validHeaders.Contains(styleName, StringComparer.CurrentCultureIgnoreCase))
            {
                break;
            }

            switch (styleName)
            {
                case "Title":
                    title = item.Text;
                    break;
                case "Author":
                    author = item.Text;
                    break;
                case "Abstract":
                    summary = item.Text;
                    break;
            }

        }

        title ??= doc.Properties.Title;
        summary ??= doc.Properties.Subject;
        author ??= doc.Properties.Creator;

        // Cleanup any oddities from the doc text.
        title = title?.ReplaceLineEndings("").Trim();
        summary = summary?.ReplaceLineEndings("").Trim();
        author = author?.ReplaceLineEndings("").Trim();

        // Use SaveDate first, then CreatedDate if unavailable.
        string date;
        if (doc.Properties.SaveDate != null)
            date = doc.Properties.SaveDate.Value.ToString("MM/dd/yyyy");
        else if (doc.Properties.CreatedDate != null)
            date = doc.Properties.CreatedDate.Value.ToString("MM/dd/yyyy");
        else
            date = DateTime.Now.ToString("MM/dd/yyyy");

        suppliedMetadata ??= new Dictionary<object, object>();

        if (!string.IsNullOrWhiteSpace(title))
            suppliedMetadata["title"] = title;
        if (!string.IsNullOrWhiteSpace(summary))
            suppliedMetadata["description"] = summary;
        if (!string.IsNullOrWhiteSpace(author))
            suppliedMetadata["author"] = author;

        suppliedMetadata["ms.date"] = date;

        string updatedFile = Path.ChangeExtension(markdownFile, ".tmp");

        await using (var writer = new StreamWriter(updatedFile))
        using (var reader = new StreamReader(markdownFile))
        {
            await writer.WriteLineAsync("---");

            await writer.WriteAsync(PersistenceUtilities.DictionaryToYamlString(suppliedMetadata));

            await writer.WriteLineAsync("---");

            bool started = !foundStart;
            while (!reader.EndOfStream)
            {
                string line = await reader.ReadLineAsync();
                if (line?.StartsWith('#') == true) started = true;
                if (started && line != null)
                    await writer.WriteLineAsync(line);
            }
        }

        File.Move(updatedFile, markdownFile, true);
    }
}