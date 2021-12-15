using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DXPlus;
using Markdig;
using Markdig.Extensions.EmphasisExtras;
using Markdig.Renderer.Docx;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using MSLearnRepos;

namespace LearnDocUtils
{
    internal sealed class LearnToDocxDXPlus : ILearnToDocx
    {
        private TripleCrownModule moduleData;
        private string markdownFile;

        public async Task Convert(ITripleCrownGitHubService tcService, string accessToken, 
            string moduleFolder, string outputFile, string zonePivot, 
            Action<string> logger, bool debug)
        {
            if (tcService == null) throw new ArgumentNullException(nameof(tcService));
            if (moduleFolder == null) throw new ArgumentNullException(nameof(moduleFolder));

            if (Directory.Exists(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' is a folder.", nameof(outputFile));

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' cannot be null or empty.", nameof(outputFile));

            if (!Path.HasExtension(outputFile))
                outputFile = Path.ChangeExtension(outputFile, "docx");

            if (File.Exists(outputFile))
                File.Delete(outputFile);

            using var document = Document.Create(outputFile);

            var rootTemp = debug ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            (moduleData, markdownFile) = await new LearnUtilities().DownloadModuleAsync(tcService, accessToken, moduleFolder, tempFolder);

            logger?.Invoke($"Converting \"{moduleData.Title}\" to {outputFile}");

            try
            {
                AddMetadata(document);
                WriteTitle(document);

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

                if (debug)
                {
                    await File.WriteAllTextAsync(Path.Combine(tempFolder, "debug.txt"), MarkdigDebug.Dump(markdownDocument));
                }

                var docWriter = new DocxObjectRenderer(document, tempFolder, zonePivot);
                docWriter.Render(markdownDocument);

                AddUnitMetadata(document);

                document.Save();
                document.Close();
            }
            finally
            {
                if (!debug)
                {
                    try
                    {
                        Directory.Delete(tempFolder, true);
                    }
                    catch
                    {
                        // ignored
                    }
                }
                else
                {
                    logger?.Invoke($"Downloaded module folder: {tempFolder}");
                }
            }
        }

        private void AddUnitMetadata(IDocument document)
        {
            var headers = document.Paragraphs
                .Where(p => p.Properties.StyleName == HeadingType.Heading1.ToString())
                .ToList();

            string user = Environment.UserName;
            if (string.IsNullOrEmpty(user))
                user = "Office User";

            foreach (var unit in moduleData.Units)
            {
                if (unit.UsesSandbox || unit.LabId != null || !string.IsNullOrEmpty(unit.InteractivityType))
                {
                    string title = unit.Title;
                    var p = headers.SingleOrDefault(p => p.Text == title);
                    p?.AttachComment(
                        document.CreateComment(user,
                            $"Sandbox: {unit.UsesSandbox}, LabId: {unit.LabId}, Interactivity: {unit.InteractivityType}"));
                }
            }
        }

        private void AddMetadata(IDocument document)
        {
            document.SetPropertyValue(DocumentPropertyName.Title, moduleData.Title);
            document.SetPropertyValue(DocumentPropertyName.Subject, moduleData.Summary);
            document.SetPropertyValue(DocumentPropertyName.Keywords, string.Join(',', moduleData.Products));
            document.SetPropertyValue(DocumentPropertyName.Comments, string.Join(',',moduleData.FriendlyLevels));
            document.SetPropertyValue(DocumentPropertyName.Category, string.Join(',', moduleData.FriendlyRoles));
            document.SetPropertyValue(DocumentPropertyName.CreatedDate, moduleData.LastUpdated.ToString("yyyy-MM-ddT00:00:00Z"));
            document.SetPropertyValue(DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);

            // Add custom data.
            document.AddCustomProperty("ModuleUid", moduleData.Uid);
            document.AddCustomProperty("MsTopic", moduleData.Metadata.MsTopic);
            document.AddCustomProperty("MsProduct", moduleData.Metadata.MsProduct);
            document.AddCustomProperty("Abstract", moduleData.Abstract);
        }

        private void WriteTitle(IDocument document)
        {
            document.AddParagraph(moduleData.Title)
                .Style(HeadingType.Title);
            document.AddParagraph($"Last modified on {moduleData.LastUpdated.ToShortDateString()} by {moduleData.Metadata.MsAuthor}@microsoft.com")
                .Style(HeadingType.Subtitle);
        }
    }
}