using System.Text;
using ConvertLearnToDoc.Shared;
using DXPlus;
using Julmar.GenMarkdown;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Extensions.Yaml;
using Markdig.Renderer.Docx;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using MSLearnRepos;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Octokit;
using YamlDotNet.Serialization;
using static System.Net.Mime.MediaTypeNames;
using Formatting = Newtonsoft.Json.Formatting;
using Paragraph = DXPlus.Paragraph;

namespace LearnDocUtils;

public static class MarkdownToDocConverter
{
    public static async Task<List<string>> ConvertMarkdownToDocx(ILearnRepoService learnRepo, string inputLocation,
        Module moduleData, string markdownFile, string docxFile, string zonePivot, bool dumpMarkdownDocument)
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
            .UseYamlFrontMatter()
            .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
            .UseMonikerRange(context)
            .UseIncludeFile(context)
            .UseQuoteSectionNote(context)
            .UseRow(context)
            .UseNestedColumn(context)
            .UseTripleColon(context)
            .UseNoloc()
            .Build();

        // Get the web location for this content
        string webLocation = moduleData != null ? moduleData.Url : "https://learn.microsoft.com/";
        if (!webLocation.EndsWith('/')) webLocation += '/';
        webLocation += inputLocation.Replace('\\', '/');
        if (!webLocation.EndsWith('/')) webLocation += '/';

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

        if (moduleData != null)
        {
            AddMetadata(moduleData, document);
        }
        else
        {
            if (markdownDocument.FirstOrDefault(b => b.GetType() == typeof(YamlFrontMatterBlock)) is YamlFrontMatterBlock header)
            {
                AddMetadata(header, document);
            }
        }

        // Write the start of the document.
        WriteTitle(document);

        // Try to get some Markdown config from the document.
        bool useAsterisksForEmphasis = markdownDocument.EnumerateBlocks()
                                           .Count(b => b is EmphasisInline { DelimiterChar: '*' })
                                       > markdownDocument.EnumerateBlocks()
                                           .Count(b => b is EmphasisInline { DelimiterChar: '_' });

        bool useAsterisksForLists = markdownDocument.EnumerateBlocks()
                                        .Count(b => b is ListBlock { BulletType: '*' })
                                    > markdownDocument.EnumerateBlocks()
                                        .Count(b => b is ListBlock { BulletType: '-' });
        SetCustomProperty(document, nameof(MarkdownFormatting.UseAsterisksForBullets), useAsterisksForLists.ToString());
        SetCustomProperty(document, nameof(MarkdownFormatting.UseAsterisksForEmphasis), useAsterisksForEmphasis.ToString());
        SetCustomProperty(document, nameof(Uri), webLocation);

        // Render the markdown tree into the Word document.
        var docWriter = new DocxObjectRenderer(document, Path.GetDirectoryName(markdownFile), 
            new DocxRendererOptions {
                Logger = text => errors.Add(text),
                ReadFile = (_, path) => GetFile(learnRepo, inputLocation, moduleData, path, Path.GetDirectoryName(markdownFile)),
                ConvertRelativeUrl = url => webLocation + url,
                ZonePivot = zonePivot
            }
        );
        docWriter.Render(markdownDocument);

        if (moduleData != null)
        {
            // Add Learn-specific unit metadata to the document - lab info, etc.
            AddUnitMetadata(docWriter, moduleData, document);
        }

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
    /// <param name="inputLocation"></param>
    /// <param name="moduleData"></param>
    /// <param name="path"></param>
    /// <param name="destinationPath"></param>
    /// <returns></returns>
    private static byte[] GetFile(ILearnRepoService learnRepo, string inputLocation, Module moduleData, string path, string destinationPath)
    {
        string remotePath;

        path = path.Replace('\\', '/').Trim();
        if (path.StartsWith('/')) // rooted path - this is relevant to the REPO folder for content.
        {
            remotePath = path[1..];
        }
        else
        {
            string checkPath = Path.Combine(destinationPath, Constants.IncludesFolder, path);
            if (File.Exists(checkPath))
                return File.ReadAllBytes(checkPath);

            remotePath = moduleData != null 
                ? Path.Combine(Path.GetDirectoryName(moduleData.Path) ?? "", Constants.IncludesFolder, path) 
                : Path.Combine(inputLocation, path);
        }

        // Try direct.
        (byte[], string) result = (null,null); 
        try
        {
            result = learnRepo.ReadFileForPathAsync(remotePath).Result;
        }
        catch (AggregateException aex)
        {
            if (aex.InnerException is not NotFoundException)
                return null;

            // If the repo is (possibly) a localized repo, look at the original (en-us) version.
            if (learnRepo.Repository.Contains('.') && learnRepo is IRemoteLearnRepoService rlrs)
            {
                var gitHubToken = Environment.GetEnvironmentVariable("GitHubToken");
                if (string.IsNullOrWhiteSpace(gitHubToken))
                    gitHubToken = null;

                string englishRepo = learnRepo.Repository[..learnRepo.Repository.IndexOf('.')];
                var enRepo = LearnRepoService.Create(rlrs.Organization, englishRepo, rlrs.Branch, gitHubToken);
                try
                {
                    result = enRepo.ReadFileForPathAsync(remotePath).Result;
                }
                catch // Not found.
                {
                    return null;
                }
            }
        }

        var (binary, text) = result;
        if (binary != null) return binary;
        if (text != null) return Encoding.Default.GetBytes(text);

        // Last check -- for local file retrieval, it's possible our "root" is the module itself.
        // This is mostly for testing, but we'll handle it here for completeness.
        if (learnRepo is not IRemoteLearnRepoService && path.Contains('/'))
        {
            string[] repoPath = learnRepo.RootPath.Replace('\\', '/').Split('/');
            string[] lfPath = remotePath.Split('/');
            string lookFor = lfPath[0];

            // Walk backwards to match up to the start of the path we're looking for.
            for (int index = repoPath.Length-1; index>0; index--)
            {
                if (string.Compare(repoPath[index], lookFor, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    remotePath = Path.Combine(
                        string.Join('/', repoPath.Take(index)),
                        string.Join('/', lfPath));
                    (binary, text) = learnRepo.ReadFileForPathAsync(remotePath).Result;
                    if (binary != null) return binary;
                    if (text != null) return Encoding.Default.GetBytes(text);
                }
            }
        }

        // Not found.
        return null;
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

    private static void SetCustomProperty(IDocument document, string name, string value)
    {
        if (!string.IsNullOrEmpty(value))
            document.CustomProperties.Add(name, value);
    }

    private static void AddMetadata(YamlFrontMatterBlock header, IDocument document)
    {
        var yaml = string.Join("\r\n",
            header.Lines.Lines.Where(s => s.ToString().Length > 0).Select(t => t.ToString()));

        var deserializer = new DeserializerBuilder().Build();
        var keys = (Dictionary<object,object>) deserializer.Deserialize<object>(yaml);

        foreach (var kvp in keys)
        {
            switch (kvp.Key.ToString()!.ToLower())
            {
                case "title":
                    document.Properties.Title = kvp.Value.ToString()!;
                    break;
                case "description":
                    document.Properties.Subject = kvp.Value.ToString()!;
                    break;
                case "author":
                    document.Properties.Creator = kvp.Value.ToString()!;
                    document.Properties.LastSavedBy = kvp.Value.ToString()!;
                    break;
                case "ms.date" 
                when DateTime.TryParse(kvp.Value.ToString()!, out var dt):
                    document.Properties.CreatedDate = dt.ToUniversalTime();
                    document.Properties.SaveDate = dt.ToUniversalTime();
                    break;
            }
        }

        if (keys.Count > 0)
        {
            var jsonText = PersistenceUtilities.ObjectToJsonString(keys);
            document.CustomProperties.Add(nameof(Metadata), jsonText);
        }
    }

    private static void AddMetadata(Module moduleData, IDocument document)
    {
        document.Properties.Title = moduleData.Title;
        document.Properties.Subject = moduleData.Summary;
        document.Properties.Creator = moduleData.Metadata?.MsAuthor ?? moduleData.Metadata?.Author ?? "TBD";
        document.Properties.LastSavedBy = moduleData.Metadata?.MsAuthor ?? moduleData.Metadata?.Author ?? "TBD";

        var jsonText = PersistenceUtilities.ObjectToJsonString(moduleData);
        SetCustomProperty(document, nameof(Module.Metadata), jsonText);

        var dt = moduleData.LastUpdated?.ToUniversalTime() ?? DateTime.UtcNow;
        document.Properties.CreatedDate = dt;
        document.Properties.SaveDate = dt;
    }

    private static void WriteTitle(IDocument document)
    {
        string title = document.Properties.Title ?? "Add Title Here";
        string author = document.Properties.Creator ?? "author";
        DateTime lastUpdated = document.Properties.CreatedDate ?? DateTime.UtcNow;

        document.Add(title.Trim('"')).Style(HeadingType.Title);
        document.Add($"Last modified on {lastUpdated.ToLocalTime().ToShortDateString()} by {author}@microsoft.com")
            .Style(HeadingType.Subtitle);
    }
}