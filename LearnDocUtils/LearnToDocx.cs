using System;
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
            string zonePivot = null, string accessToken = null, bool debug = false)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(url);
            await ConvertFromRepoAsync(repo, branch, folder, outputFile, zonePivot, accessToken, debug);
        }

        public static async Task ConvertFromRepoAsync(string repo, string branch, string folder,
                string outputFile, string zonePivot = null, string accessToken = null, bool debug = false)
        {
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            accessToken = string.IsNullOrEmpty(accessToken)
                ? GithubHelper.ReadDefaultSecurityToken()
                : accessToken;

            await Convert(
                TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken),
                accessToken, folder, outputFile, zonePivot, debug);
        }

        public static async Task ConvertFromFolderAsync(string learnFolder, string outputFile, string zonePivot = null, bool debug = false)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            await Convert(
                TripleCrownGitHubService.CreateLocal(learnFolder), null, learnFolder, outputFile, zonePivot, debug);
        }

        private static async Task Convert(ITripleCrownGitHubService tcService,
            string accessToken, string moduleFolder, string docxFile,
            string zonePivot, bool debug)
        {
            var rootTemp = debug ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            // Download the module
            var (module, markdownFile) = await new LearnUtilities().DownloadModuleAsync(tcService, accessToken, moduleFolder, tempFolder);

            try
            {
                // Convert the file.
                await ConvertMarkdownToDocx(module, markdownFile, docxFile, zonePivot);
            }
            finally
            {
                if (!debug)
                {
                    Directory.Delete(tempFolder, true);
                }
            }
        }

        private static async Task ConvertMarkdownToDocx(TripleCrownModule moduleData, string markdownFile, string docxFile, string zonePivot)
        {
            using var document = Document.Create(docxFile);
            AddMetadata(moduleData, document);
            WriteTitle(moduleData, document);

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

            string markdownText = await File.ReadAllTextAsync(markdownFile);
            var markdownDocument = Markdown.Parse(markdownText, pipeline);

            var docWriter = new DocxObjectRenderer(document, Path.GetDirectoryName(markdownFile), zonePivot);
            docWriter.Render(markdownDocument);

            AddUnitMetadata(moduleData, document);

            document.Save();
            document.Close();
        }

        private static void AddUnitMetadata(TripleCrownModule moduleData, IDocument document)
        {
            var headers = document.Paragraphs
                .Where(p => p.Properties.StyleName == HeadingType.Heading1.ToString())
                .ToList();

            string user = Environment.UserName;
            if (string.IsNullOrEmpty(user))
                user = "Office User";

            foreach (var unit in moduleData.Units)
            {
                if (unit.UsesSandbox)
                {
                    string title = unit.Title;
                    var p = headers.SingleOrDefault(p => p.Text == title);
                    string commentText = "sandbox";
                    if (!string.IsNullOrEmpty(unit.InteractivityType))
                        commentText += $" interactivity:{unit.InteractivityType}";
                    if (!string.IsNullOrEmpty(unit.Notebook))
                        commentText += $" notebook:{unit.Notebook.Trim()}";
                    p?.AttachComment(document.CreateComment(user, commentText));
                }
                else if (unit.LabId != null)
                {
                    string title = unit.Title;
                    var p = headers.SingleOrDefault(p => p.Text == title);
                    p?.AttachComment(document.CreateComment(user, $"labId:{unit.LabId}"));
                }
                else if (!string.IsNullOrEmpty(unit.InteractivityType))
                {
                    string title = unit.Title;
                    var p = headers.SingleOrDefault(p => p.Text == title);
                    p?.AttachComment(document.CreateComment(user, $"{unit.InteractivityType}"));
                }
            }
        }

        private static void AddMetadata(TripleCrownModule moduleData, IDocument document)
        {
            document.SetPropertyValue(DocumentPropertyName.Title, moduleData.Title);
            document.SetPropertyValue(DocumentPropertyName.Subject, moduleData.Summary);
            document.SetPropertyValue(DocumentPropertyName.Keywords, string.Join(',', moduleData.Products));
            document.SetPropertyValue(DocumentPropertyName.Comments, string.Join(',', moduleData.FriendlyLevels));
            document.SetPropertyValue(DocumentPropertyName.Category, string.Join(',', moduleData.FriendlyRoles));
            document.SetPropertyValue(DocumentPropertyName.CreatedDate, moduleData.LastUpdated.ToString("yyyy-MM-ddT00:00:00Z"));
            document.SetPropertyValue(DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);

            // Add custom data.
            document.AddCustomProperty("ModuleUid", moduleData.Uid);
            document.AddCustomProperty("MsTopic", moduleData.Metadata.MsTopic);
            document.AddCustomProperty("MsProduct", moduleData.Metadata.MsProduct);
            document.AddCustomProperty("Abstract", moduleData.Abstract);
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