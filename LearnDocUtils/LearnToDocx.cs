using MSLearnRepos;
using System.Text;
using DXPlus;
using Julmar.GenMarkdown;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Renderer.Docx;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Formatting = Newtonsoft.Json.Formatting;
using Paragraph = DXPlus.Paragraph;

namespace LearnDocUtils;

public static class LearnToDocx
{
    public static async Task<List<string>> ConvertFromUrlAsync(string url, string outputFile, 
                                            string accessToken = null, DocumentOptions options = null)
    {
        if (string.IsNullOrEmpty(url))
            throw new ArgumentNullException(nameof(url));

        var metadata = await DocsMetadata.LoadFromUrlAsync(url);
        return await ConvertFromRepoAsync(metadata.Organization, metadata.Repository, metadata.Branch, 
            Path.GetDirectoryName(metadata.ContentPath), outputFile, accessToken, options);
    }

    public static async Task<List<string>> ConvertFromRepoAsync(string organization, string repo, string branch, string folder,
        string outputFile, string accessToken = null, DocumentOptions options = null)
    {
        if (string.IsNullOrEmpty(organization)) 
            throw new ArgumentException($"'{nameof(organization)}' cannot be null or empty.", nameof(organization));
        if (string.IsNullOrEmpty(repo))
            throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
        if (string.IsNullOrEmpty(folder))
            throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

        return await Convert(LearnRepoService.Create(organization, repo, branch, accessToken), 
            folder, outputFile, options);
    }

    public static async Task<List<string>> ConvertFromFolderAsync(string learnFolder, string outputFile, DocumentOptions options = null)
    {
        if (string.IsNullOrWhiteSpace(learnFolder))
            throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

        if (!Directory.Exists(learnFolder))
            throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

        return await Convert(LearnRepoService.Create(learnFolder), learnFolder, outputFile, options);
    }

    private static async Task<List<string>> Convert(ILearnRepoService learnRepo,
        string moduleFolder, string docxFile, DocumentOptions options)
    {
        var rootTemp = options?.Debug == true ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
        var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
        while (Directory.Exists(tempFolder))
        {
            tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
        }

        // Download the module
        var (module, markdownFile) = await ModuleCombiner.DownloadModuleAsync(
            learnRepo, moduleFolder, tempFolder,
            options?.EmbedNotebookContent == true);

        try
        {
            // Convert the file.
            return await ConvertMarkdownToDocx(learnRepo, module, markdownFile, docxFile, options?.ZonePivot, options?.Debug==true);
        }
        finally
        {
            if (options is {Debug: false})
            {
                Directory.Delete(tempFolder, true);
            }
        }
    }


    private static async Task<List<string>> ConvertMarkdownToDocx(ILearnRepoService learnRepo, Module moduleData, 
        string markdownFile, string docxFile, string zonePivot, bool dumpMarkdownDocument)
    {
        var errors = new List<string>();

        MarkdownContext.LogActionDelegate Log(string level) 
            => (code, message, origin, line) => errors.Add($"{level}: {code} - {message}");

        var context = new MarkdownContext(
            logInfo: (a, b, c, d) => { },
            logSuggestion: Log("suggestion"),
            logWarning: Log("warning"),
            logError: Log("error"));

        var pipelineBuilder = new MarkdownPipelineBuilder();
        var pipeline = pipelineBuilder
            .UseGridTables()
            .UseMediaLinks()
            .UsePipeTables()
            .UseListExtras()
            .UseTaskLists()
            .UseAutoLinks()
            .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
            .UseIncludeFile(context)
            .UseQuoteSectionNote(context)
            .UseRow(context)
            .UseNestedColumn(context)
            .UseTripleColon(context)
            .UseNoloc()
            .Build();

        string markdownText = await File.ReadAllTextAsync(markdownFile);

        // Pre-process the :::image tag .. we're using the public version (v2) which has problems
        // with end tags when they aren't on a separate line.
        markdownText = FixImageExtension(markdownText);

        // Parse the Markdown tree
        var markdownDocument = Markdown.Parse(markdownText, pipeline);

        // Dump the markdown contents.
        if (dumpMarkdownDocument)
        {
            var text = MarkdigExtensions.Dump(markdownDocument);
            var folder = Path.GetDirectoryName(docxFile) ?? "";
            await File.WriteAllTextAsync(Path.Combine(folder, "markdown-tree.g.txt"), text);
        }

        // Create the Word document
        using var document = Document.Create(docxFile);
        AddMetadata(moduleData, document);
        WriteTitle(moduleData, document);


        // Try to get some Markdown config from the document.
        bool useAsterisksForEmphasis = markdownDocument.EnumerateBlocks()
                                           .Count(b => b is EmphasisInline { DelimiterChar: '*' })
                                       > markdownDocument.EnumerateBlocks()
                                           .Count(b => b is EmphasisInline { DelimiterChar: '_' });

        bool useAsterisksForLists = markdownDocument.EnumerateBlocks()
                                        .Count(b => b is Markdig.Syntax.ListBlock { BulletType: '*' })
                                    > markdownDocument.EnumerateBlocks()
                                        .Count(b => b is Markdig.Syntax.ListBlock { BulletType: '-' });
        SetCustomProperty(document, nameof(MarkdownFormatting.UseAsterisksForBullets), useAsterisksForLists.ToString());
        SetCustomProperty(document, nameof(MarkdownFormatting.UseAsterisksForEmphasis), useAsterisksForEmphasis.ToString());

        // Render the markdown tree into the Word document.
        var docWriter = new DocxObjectRenderer(document, Path.GetDirectoryName(markdownFile), 
            new DocxRendererOptions {
                Logger = text => errors.Add(text),
                ReadFile = (_, path) => GetFile(learnRepo, moduleData, path, Path.GetDirectoryName(markdownFile)),
                ZonePivot = zonePivot
            }
        );
        docWriter.Render(markdownDocument);

        // Add Learn-specific unit metadata to the document - lab info, etc.
        AddUnitMetadata(docWriter, moduleData, document);

        document.Save();
        document.Close();

        return errors;
    }

    /// <summary>
    /// Used to retrieve a file when it wasn't part of the original repo folder fetch.
    /// This happens when global files outside the Learn module folder are referenced, or
    /// when a file isn't in a traditional location such as "media".
    /// </summary>
    /// <param name="learnRepo"></param>
    /// <param name="moduleData"></param>
    /// <param name="path"></param>
    /// <param name="destinationPath"></param>
    /// <returns></returns>
    private static byte[] GetFile(ILearnRepoService learnRepo, Module moduleData, string path, string destinationPath)
    {
        string checkPath = Path.Combine(destinationPath, Constants.IncludesFolder, path);
        if (File.Exists(checkPath))
            return File.ReadAllBytes(checkPath);

        string remotePath = Path.Combine(Path.GetDirectoryName(moduleData.Path), Constants.IncludesFolder, path);
        var (binary, text) = learnRepo.ReadFileForPathAsync(remotePath).Result;
        if (binary != null) 
            return binary;
        return text != null 
            ? Encoding.Default.GetBytes(text) 
            : null;
    }

    private static string FixImageExtension(string text)
    {
        const string marker = ":::image-end:::";

        int index = text.IndexOf(marker, StringComparison.InvariantCultureIgnoreCase);
        while (index > 0)
        {
            if (text[index - 1] != '\r' && text[index - 1] != '\n')
                text = text.Insert(index, Environment.NewLine);

            // Backup and find the end marker -- we need to make sure the description text is bounded by CRLF for
            // the existing v2 extension to pick it up.

            int start = index - 3;
            for (; start > 0; start--)
            {
                if (text.Substring(start, 3) == ":::")
                    break;
            }

            if (start > 0)
            {
                start += 3;
                while (text[start] == ' ') start++;
                if (text[start] != '\r' && text[start] != '\n')
                    text = text.Insert(start, Environment.NewLine);
            }

            index = text.IndexOf(marker, index+marker.Length, StringComparison.InvariantCultureIgnoreCase);
        }

        return text;
    }

    private static void AddUnitMetadata(IDocxRenderer renderer, Module moduleData, IDocument document)
    {
        var headers = document.Paragraphs
            .Where(p => p.Properties.StyleName == HeadingType.Heading1.ToString())
            .ToList();

        string user = Environment.UserInteractive ? Environment.UserName : Environment.GetEnvironmentVariable("CommentUserName");
        if (string.IsNullOrEmpty(user))
            user = "Office User";

        foreach (var unit in moduleData.Units)
        {
            var unitHeaderParagraph = FindParagraphByTitle(headers, moduleData.Units, unit);
            if (unitHeaderParagraph == null) continue;

            string commentText = null;
            if (unit.UsesSandbox)
            {
                commentText = "sandbox";
                if (!string.IsNullOrEmpty(unit.InteractivityType))
                    commentText += $" interactivity:{unit.InteractivityType}";
                if (!string.IsNullOrEmpty(unit.Notebook))
                    commentText += $" notebook:{unit.Notebook.Trim()}";
            }
            else if (unit.LabId != null)
            {
                commentText = $"labId:{unit.LabId}";
            }
            else if (!string.IsNullOrEmpty(unit.InteractivityType))
            {
                commentText = $"interactivity:{unit.InteractivityType}";
            }

            if (commentText != null)
            {
                renderer.AddComment(unitHeaderParagraph, commentText);
            }
        }
    }

    private static Paragraph FindParagraphByTitle(IEnumerable<Paragraph> headers, IEnumerable<ModuleUnit> moduleDataUnits, ModuleUnit unit)
    {
        // Multiple units can have the same title .. not a good practice, but it happens.
        // Find the specific title we are looking for by index.
        string title = unit.Title;
        int pos = moduleDataUnits.Where(u => u.Title == title).ToList().IndexOf(unit);
        return headers.Where(p => p.Text == title).Skip(pos).FirstOrDefault();
    }

    private static void SetProperty(IDocument document, DocumentPropertyName name, string value)
    {
        if (!string.IsNullOrEmpty(value))
            document.SetPropertyValue(name, value);
    }

    private static void SetCustomProperty(IDocument document, string name, string value)
    {
        if (!string.IsNullOrEmpty(value))
            document.AddCustomProperty(name, value);
    }

    private static void AddMetadata(Module moduleData, IDocument document)
    {
        SetProperty(document, DocumentPropertyName.Title, moduleData.Title);
        SetProperty(document, DocumentPropertyName.Subject, moduleData.Summary);
        SetProperty(document, DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);
        SetProperty(document, DocumentPropertyName.LastSavedBy, moduleData.Metadata.MsAuthor);

        SetProperty(document, DocumentPropertyName.Comments, moduleData.Uid);
        SetCustomProperty(document, nameof(Module.Metadata), 
            JsonConvert.SerializeObject(moduleData, Formatting.None,
                                                new JsonSerializerSettings {
                                                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                                                    NullValueHandling = NullValueHandling.Ignore
                                                }));

        var dt = (moduleData.LastUpdated?.ToUniversalTime() ?? DateTime.UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ");
        SetProperty(document, DocumentPropertyName.CreatedDate, dt);
        SetProperty(document, DocumentPropertyName.SaveDate, dt);
    }

    private static void WriteTitle(Module moduleData, IDocument document)
    {
        document.AddParagraph(moduleData.Title)
            .Style(HeadingType.Title);
        document.AddParagraph($"Last modified on {(moduleData.LastUpdated ?? DateTime.Now).ToShortDateString()} by {moduleData.Metadata.MsAuthor}@microsoft.com")
            .Style(HeadingType.Subtitle);
    }
}