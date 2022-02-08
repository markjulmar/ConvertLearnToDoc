using System;
using System.Collections.Generic;
using MSLearnRepos;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DXPlus;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Renderer.Docx;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace LearnDocUtils
{
    public static class LearnToDocx
    {
        public static async Task ConvertFromUrlAsync(string url, string outputFile, 
            string zonePivot = null, string accessToken = null, DocumentOptions options = null)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnResolver.LocationFromUrlAsync(url);
            await ConvertFromRepoAsync(repo, branch, folder, outputFile, zonePivot, accessToken, options);
        }

        public static async Task ConvertFromRepoAsync(string repo, string branch, string folder,
                string outputFile, string zonePivot = null, string accessToken = null, DocumentOptions options = null)
        {
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            accessToken = string.IsNullOrEmpty(accessToken) ? GithubHelper.ReadDefaultSecurityToken() : accessToken;
            await Convert(TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken), folder, outputFile, zonePivot, options);
        }

        public static async Task ConvertFromFolderAsync(string learnFolder, string outputFile, string zonePivot = null, DocumentOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            await Convert(TripleCrownGitHubService.CreateLocal(learnFolder), learnFolder, outputFile, zonePivot, options);
        }

        private static async Task Convert(ITripleCrownGitHubService tcService,
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
                await ConvertMarkdownToDocx(module, markdownFile, docxFile, zonePivot, options?.Debug==true);
            }
            finally
            {
                if (options is {Debug: false})
                {
                    Directory.Delete(tempFolder, true);
                }
            }
        }

        private static async Task ConvertMarkdownToDocx(TripleCrownModule moduleData, string markdownFile, string docxFile, string zonePivot, bool dumpMarkdownDocument)
        {
            var context = new MarkdownContext();
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

            // Parse the Markdown tree
            string markdownText = await File.ReadAllTextAsync(markdownFile);
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

            // Render the markdown tree into the Word document.
            var docWriter = new DocxObjectRenderer(document, Path.GetDirectoryName(markdownFile), zonePivot);
            docWriter.Render(markdownDocument);

            // Add Learn-specific unit metadata to the document - lab info, etc.
            AddUnitMetadata(docWriter, moduleData, document);

            document.Save();
            document.Close();
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

            var dt = (moduleData.LastUpdated == default ? DateTime.Now : moduleData.LastUpdated).ToString("yyyy-MM-ddT00:00:00Z");
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