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
    sealed class MarkdownToDocxDxPlus : IMarkdownToDocx
    {
        public async Task Convert(TripleCrownModule moduleData, string markdownFile, string docxFile, string zonePivot)
        {
            using var document = Document.Create(docxFile);
            AddMetadata(moduleData, document);
            WriteTitle(moduleData,document);

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
