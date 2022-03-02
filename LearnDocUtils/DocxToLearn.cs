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
                    options.UseAsterisksForBullets = yesNo.ToString() == "True";
                if (doc.CustomProperties.TryGetValue(nameof(MarkdownOptions.UseAsterisksForEmphasis), out yesNo))
                    options.UseAsterisksForEmphasis = yesNo.ToString() == "True";
            }

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
            if (!options.Debug)
            {
                if (File.Exists(markdownFile))
                    File.Delete(markdownFile);
            }
        }
    }

    /// <summary>
    /// Do some post-conversion cleanup of markers, paths, and triple-colon placeholders.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    private static string PostProcessMarkdown(string text)
    {
        text = text.Trim('\r').Trim('\n');
        text = text.Replace("\xa0", " "); // TODO: should this be &nbsp;?

        text = Regex.Replace(text, @"{rgn (.*?)}", m => $"<rgn>{m.Groups[1].Value.Trim()}</rgn>");
        text = Regex.Replace(text, @"{zonePivot:(.*?)}", m => $":::zone pivot={m.Groups[1].Value.Trim()}");
        text = Regex.Replace(text, @"{end-zonePivot:(.*?)}", m => $":::zone-end");
        text = Regex.Replace(text, @"{include ""(.*?)"".*}", m => $"[!include[]({m.Groups[1].Value.Trim()})]");

        // WWL templates use prefixes, convert these to our quote block note types.
        text = ProcessNotes(text, "note: ", "NOTE");
        text = ProcessNotes(text, "(*) ", "TIP");

        return text;
    }

    private static string ProcessNotes(string text, string lookFor, string header)
    {
        int index = text.IndexOf(lookFor, StringComparison.CurrentCultureIgnoreCase);
        while (index > 0)
        {
            if (text[index - 1] == '\r' || text[index - 1] == '\n')
            {
                int end = text.IndexOf('\n', index);
                if (end > index)
                {
                    end++;
                    string noteBlock = text.Substring(index + lookFor.Length, end - index - lookFor.Length).TrimEnd('\r', '\n');
                    text = text.Remove(index, end - index)
                        .Insert(index, $"> [!{header.ToUpper()}]{Environment.NewLine}> {noteBlock}{Environment.NewLine}");
                }
            }

            index = text.IndexOf(lookFor, index+1, StringComparison.CurrentCultureIgnoreCase);
        }

        return text;
    }
}