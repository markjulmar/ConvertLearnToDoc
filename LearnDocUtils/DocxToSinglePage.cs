using Docx.Renderer.Markdown;
using DXPlus;
using MSLearnRepos;

namespace LearnDocUtils;

public static class DocxToSinglePage
{
    public static async Task ConvertAsync(string docxFile, string markdownFile, MarkdownOptions options)
    {
        if (string.IsNullOrWhiteSpace(docxFile))
            throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
        if (!File.Exists(docxFile))
            throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
        if (string.IsNullOrWhiteSpace(markdownFile))
            throw new ArgumentException($"'{nameof(markdownFile)}' cannot be null or whitespace.", nameof(markdownFile));

        options ??= new MarkdownOptions();

        string baseFilename = Path.GetFileNameWithoutExtension(docxFile);
        string outputFolder = Path.GetDirectoryName(markdownFile) ?? Directory.GetCurrentDirectory();
        string mediaFolder = Path.Combine(outputFolder, Constants.MediaFolder);

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

        // Do some common post-processing.
        var markdownText = DocToMarkdownRenderer.PostProcessMarkdown(await File.ReadAllTextAsync(markdownFile));
        await File.WriteAllTextAsync(markdownFile, markdownText);

        // Pull out and write the YAML header
        await WriteMetadata(docxFile, markdownFile);
    }

    private static async Task WriteMetadata(string docxFile, string markdownFile)
    {
        var doc = Document.Load(docxFile);

        string title = null, author = null, summary = null;

        foreach (var item in doc.Paragraphs)
        {
            var styleName = item.Properties.StyleName;
            if (styleName == "Heading1") break;
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

        // Use SaveDate first, then CreatedDate if unavailable.
        string date;
        if (doc.Properties.SaveDate != null)
            date = doc.Properties.SaveDate.Value.ToString("MM/dd/yyyy");
        else if (doc.Properties.CreatedDate != null)
            date = doc.Properties.CreatedDate.Value.ToString("MM/dd/yyyy");
        else
            date = DateTime.Now.ToString("MM/dd/yyyy");

        string additionalMetadata = doc.CustomProperties.TryGetValue(nameof(Metadata), out var property) == true
            ? property?.Value
            : null;

        string updatedFile = Path.ChangeExtension(markdownFile, ".tmp");

        await using (var writer = new StreamWriter(updatedFile))
        using (var reader = new StreamReader(markdownFile))
        {
            await writer.WriteLineAsync("---");
            await writer.WriteLineAsync($"title: {title}");
            await writer.WriteLineAsync($"description: {summary}");
            await writer.WriteLineAsync($"author: {author}");
            await writer.WriteLineAsync($"ms.date: {date}");
            if (additionalMetadata != null)
                await writer.WriteAsync(additionalMetadata);
            await writer.WriteLineAsync("---");

            bool started = false;
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