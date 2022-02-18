using System;
using System.Collections.Generic;
using MSLearnRepos;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DXPlus;
using Julmar.GenMarkdown;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Renderer.Docx;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using Paragraph = DXPlus.Paragraph;

namespace LearnDocUtils
{
    public static class LearnToDocx
    {
        public static async Task<List<string>> ConvertFromUrlAsync(string url, string outputFile, 
            string zonePivot = null, string accessToken = null, DocumentOptions options = null)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnResolver.LocationFromUrlAsync(url);
            return await ConvertFromRepoAsync(repo, branch, folder, outputFile, zonePivot, accessToken, options);
        }

        public static async Task<List<string>> ConvertFromRepoAsync(string repo, string branch, string folder,
                string outputFile, string zonePivot = null, string accessToken = null, DocumentOptions options = null)
        {
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            accessToken = string.IsNullOrEmpty(accessToken) ? GithubHelper.ReadDefaultSecurityToken() : accessToken;
            return await Convert(TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken), folder, outputFile, zonePivot, options);
        }

        public static async Task<List<string>> ConvertFromFolderAsync(string learnFolder, string outputFile, string zonePivot = null, DocumentOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            return await Convert(TripleCrownGitHubService.CreateLocal(learnFolder), learnFolder, outputFile, zonePivot, options);
        }

        private static async Task<List<string>> Convert(ITripleCrownGitHubService tcService,
            string moduleFolder, string docxFile,
            string zonePivot, DocumentOptions options)
        {
            var rootTemp = options?.Debug == true ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            // Download the module
            var (module, markdownFile) = await new ModuleDownloader().DownloadModuleAsync(
                    tcService, moduleFolder, tempFolder,
                    options?.EmbedNotebookContent == true);

            try
            {
                // Convert the file.
                return await ConvertMarkdownToDocx(module, markdownFile, docxFile, zonePivot, options?.Debug==true);
            }
            finally
            {
                if (options is {Debug: false})
                {
                    Directory.Delete(tempFolder, true);
                }
            }
        }


        private static async Task<List<string>> ConvertMarkdownToDocx(TripleCrownModule moduleData, string markdownFile, string docxFile, string zonePivot, bool dumpMarkdownDocument)
        {
            var errors = new List<string>();

            MarkdownContext.LogActionDelegate Log(string level) 
                => (code, message, origin, line) => errors.Add($"{level}: {code} - {message}");

            var context = new MarkdownContext(
                logInfo: (a, b, c, d) => { },
                logSuggestion: Log("suggestion"),
                logWarning: Log("warning"),
                logError: Log("error"));
            // TODO: add readFile support?

            var pipelineBuilder = new MarkdownPipelineBuilder();
            var pipeline = pipelineBuilder
                .UseAbbreviations()
                .UseAutoIdentifiers()
                //.UseCitations()
                //.UseCustomContainers()
                //.UseDefinitionLists()
                //.UseFigures()
                //.UseFooters()
                //.UseFootnotes()
                .UseGridTables()
                .UseMathematics()
                .UseMediaLinks()
                .UsePipeTables()
                .UseListExtras()
                .UseTaskLists()
                //.UseDiagrams()
                .UseAutoLinks()
                .UseEmphasisExtras(EmphasisExtraOptions.Strikethrough)
                .UseIncludeFile(context)
                .UseQuoteSectionNote(context)
                .UseRow(context)
                .UseNestedColumn(context)
                .UseTripleColon(context)
                .UseNoloc()
                .UseGenericAttributes() // Must be last as it is one parser that is modifying other parsers
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
            var docWriter = new DocxObjectRenderer(document, Path.GetDirectoryName(markdownFile), zonePivot);
            docWriter.Render(markdownDocument);

            // Add Learn-specific unit metadata to the document - lab info, etc.
            AddUnitMetadata(docWriter, moduleData, document);

            document.Save();
            document.Close();

            return errors;
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

        private static void AddUnitMetadata(IDocxRenderer renderer, TripleCrownModule moduleData, IDocument document)
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

        private static Paragraph FindParagraphByTitle(IEnumerable<Paragraph> headers, IEnumerable<TripleCrownUnit> moduleDataUnits, TripleCrownUnit unit)
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

        private static void AddMetadata(TripleCrownModule moduleData, IDocument document)
        {
            SetProperty(document, DocumentPropertyName.Title, moduleData.Title);
            SetProperty(document, DocumentPropertyName.Subject, moduleData.Summary);
            SetProperty(document, DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);
            SetProperty(document, DocumentPropertyName.LastSavedBy, moduleData.Metadata.MsAuthor);

            SetProperty(document, DocumentPropertyName.Comments, moduleData.Uid);
            SetCustomProperty(document, nameof(TripleCrownModule.Metadata), JsonHelper.ToJson(moduleData));

            var dt = (moduleData.LastUpdated == default ? DateTime.UtcNow : moduleData.LastUpdated.ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ssZ");
            SetProperty(document, DocumentPropertyName.CreatedDate, dt);
            SetProperty(document, DocumentPropertyName.SaveDate, dt);
        }

        private static void WriteTitle(TripleCrownModule moduleData, IDocument document)
        {
            document.AddParagraph(moduleData.Title)
                .Style(HeadingType.Title);
            document.AddParagraph($"Last modified on {moduleData.LastUpdated.ToShortDateString()} by {moduleData.Metadata.MsAuthor}@microsoft.com")
                .Style(HeadingType.Subtitle);
        }
    }
}