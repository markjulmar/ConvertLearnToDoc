#define MANUAL_CHECK
using ConvertLearnToDoc.AzureFunctions.Models;
using Microsoft.Extensions.Logging;
using System.IO;
using Xunit;
using Shouldly;
using System.Threading.Tasks;
using LearnDocUtils;
using MSLearnRepos;
using System;
using System.IO.Compression;
using System.Linq;
using MSLearnRepos.Internal;
using Xunit.Abstractions;

namespace ConvertLearnToDoc.Tests.Int;

public class IntegrationTests
{
    private const string Repo = "learn-pr";
    private const string Branch = "master";
    private const string IncludeFolder = "includes";
    private const string Module = "route-and-process-data-logic-apps";

    private readonly string repoFolder = $"/{Repo}/azure/{Module}";

    private readonly string? gitHubToken;
    private readonly ILogger logger = TestFactory.CreateLogger();
    private readonly ITestOutputHelper output;

    public IntegrationTests(ITestOutputHelper output)
    {
        // Allow for local GitHub token (stored in My Documents/github-token.txt or ~/github-token.txt).
        Environment.SetEnvironmentVariable("AZURE_FUNCTIONS_ENVIRONMENT", "Development");
        gitHubToken = Environment.GetEnvironmentVariable("GitHubToken");
        this.output = output;
    }

    [Fact]
    public async Task LearnToDoc_should_create_docx_file()
    {
        var doc = new LearnToDocModel()
        {
            Branch = Branch,
            EmbedNotebookData = false,
            Folder = repoFolder,
            Repository = Repo,
            ZonePivot = ""
        };

        var fileContents = await TestUtilities.CreateWordDoc(doc, logger);
        fileContents.ShouldNotBeNull();
    }

    [Fact]
    public async Task DocToLearn_should_create_markdown_files()
    {
        var fileName = Path.ChangeExtension(Module, "docx");
        var wordDocFilePath = Path.Combine(Directory.GetCurrentDirectory(), @$"..\..\..\{fileName}");

        var fileContents = await TestUtilities.ConvertDocxToModuleZipFile(wordDocFilePath, fileName, logger);
        fileContents.ShouldNotBeNull();
    }

    [Fact]
    public async Task Should_roundTrip()
    {
#if MANUAL_CHECK
        var outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "LearnToDocTests");
#else
        var outputPath = Path.Combine(Path.GetTempPath(), "LearnToDocTests");
#endif

        //Generate and Download Word Doc from existing module
        var doc = new LearnToDocModel()
        {
            Branch = Branch,
            EmbedNotebookData = false,
            Folder = repoFolder,
            Repository = Repo,
            ZonePivot = ""
        };

        try
        {

            var wordDocFileContents = await TestUtilities.CreateWordDoc(doc, logger);
            var wordDocTempFilePath = Path.Combine(outputPath, Path.ChangeExtension(Path.GetTempFileName(), "docx"));
            await File.WriteAllBytesAsync(wordDocTempFilePath, wordDocFileContents);

            //Generate markdown from Word doc
            var fileName = Path.ChangeExtension(Module, "docx");
            var zipFileContents = await TestUtilities.ConvertDocxToModuleZipFile(wordDocTempFilePath, fileName, logger);
            var zipFileTempFilePath = Path.Combine(outputPath, Path.GetTempFileName());
            zipFileTempFilePath = Path.ChangeExtension(zipFileTempFilePath, ".zip");
            await File.WriteAllBytesAsync(zipFileTempFilePath, zipFileContents);
            var unzippedPath = Path.Combine(outputPath, "ConvertedMD");
            if (Directory.Exists(unzippedPath))
            {
                Directory.Delete(unzippedPath, true);
            }

            ZipFile.ExtractToDirectory(zipFileTempFilePath, unzippedPath, overwriteFiles: true);

            //Download original source files
            var originalPath = Path.Combine(outputPath, "OriginalMD");
            var tcService = TripleCrownGitHubService.CreateFromToken(Repo, Branch, gitHubToken);
            await new ModuleDownloader().DownloadModuleAsync(tcService,
                learnFolder: repoFolder,
                outputFolder: originalPath,
                embedNotebooks: false);

            foreach (var originalFilePath in Directory.GetFiles(originalPath, "*.yml", SearchOption.TopDirectoryOnly))
            {
                var originalFileName = Path.GetFileName(originalFilePath);
                if (File.Exists(Path.Combine(unzippedPath, originalFileName)))
                {
                    var originalText = await File.ReadAllTextAsync(Path.Combine(originalPath, originalFileName));
                    var newText = await File.ReadAllTextAsync(Path.Combine(unzippedPath, originalFileName));

                    var type = TripleCrownParser.DetermineTypeFromHeader(originalText);
                    switch (type)
                    {
                        case YamlType.Module:
                            originalFileName.ShouldBe("index.yml");
                            CompareModuleIndexes(
                                TripleCrownParser.LoadContentFromString<TripleCrownModule>(type, originalText),
                                TripleCrownParser.LoadContentFromString<TripleCrownModule>(type, newText));
                            break;
                        case YamlType.Unit:
                            CompareModuleUnits(
                                TripleCrownParser.LoadContentFromString<TripleCrownUnit>(type, originalText),
                                TripleCrownParser.LoadContentFromString<TripleCrownUnit>(type, newText));
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    Assert.Fail($"Expected YAML file \"{originalFileName}\" missing from generated content");
                }
            }

            // Do the same for Markdown content
            originalPath = Path.Combine(originalPath, IncludeFolder);
            unzippedPath = Path.Combine(unzippedPath, IncludeFolder);

            foreach (var originalFilePath in Directory.GetFiles(originalPath, "*.md", SearchOption.TopDirectoryOnly))
            {
                var originalFileName = Path.GetFileName(originalFilePath);
                if (File.Exists(Path.Combine(unzippedPath, originalFileName)))
                {
                    string fn1 = Path.Combine(originalPath, originalFileName);
                    string fn2 = Path.Combine(unzippedPath, originalFileName);

                    var diff = MarkdownComparer.CompareFiles(fn1, fn2).ToList();
                    if (diff.Count > 0)
                    {
                        // TODO: beef this up so we can assert.
                        output.WriteLine($"Found differences in Markdown content {originalFileName}:");
                        diff.ForEach(d => output.WriteLine(d.ToString()));
                    }
                }
                else
                {
                    Assert.Fail($"Expected Markdown file \"{originalFileName}\" missing from generated content");
                }
            }

        }
        finally
        {
#if !MANUAL_CHECK
            Directory.Delete(outputPath, true);
#endif
        }
    }

    private static void CompareModuleUnits(TripleCrownUnit originalYamlDoc, TripleCrownUnit newYamlDoc)
    {
        originalYamlDoc.ShouldNotBeNull();
        originalYamlDoc.Metadata.ShouldNotBeNull();

        newYamlDoc.ShouldNotBeNull();
        newYamlDoc.Metadata.ShouldNotBeNull();

        newYamlDoc.Uid.ShouldBeEquivalentTo(originalYamlDoc.Uid);
        newYamlDoc.Title.ShouldBeEquivalentTo(originalYamlDoc.Title);

        newYamlDoc.UsesSandbox.ShouldBeEquivalentTo(originalYamlDoc.UsesSandbox);
        newYamlDoc.LabId.ShouldBeEquivalentTo(originalYamlDoc.LabId);
        newYamlDoc.Content.ShouldBeEquivalentTo(originalYamlDoc.Content);
        newYamlDoc.DurationInMinutes.ShouldBeEquivalentTo(originalYamlDoc.DurationInMinutes);
        newYamlDoc.InteractivityType.ShouldBeEquivalentTo(originalYamlDoc.InteractivityType);
        newYamlDoc.Notebook.ShouldBeEquivalentTo(originalYamlDoc.Notebook);

        // Default value provided by system - can be set in produced content.
        if (originalYamlDoc.Quiz is {Title: null})
            originalYamlDoc.Quiz.Title = "Knowledge check";

        newYamlDoc.Quiz.ShouldBeEquivalentTo(originalYamlDoc.Quiz);
        newYamlDoc.Tasks.ShouldBeEquivalentTo(originalYamlDoc.Tasks);

        newYamlDoc.Metadata.Title.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Title);
        newYamlDoc.Metadata.Description.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Description);
        newYamlDoc.Metadata.Author.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Author);
        newYamlDoc.Metadata.MsAuthor.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsAuthor);
        //newYamlDoc.Metadata.MsDate.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsDate);
        newYamlDoc.Metadata.MsProduct.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsProduct);
        newYamlDoc.Metadata.MsTopic.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsTopic);
        //newYamlDoc.Metadata.Robots.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Robots);
    }

    private static void CompareModuleIndexes(TripleCrownModule originalYamlDoc, TripleCrownModule newYamlDoc)
    {
        originalYamlDoc.ShouldNotBeNull();
        originalYamlDoc.Metadata.ShouldNotBeNull();

        newYamlDoc.ShouldNotBeNull();
        newYamlDoc.Metadata.ShouldNotBeNull();

        newYamlDoc.Uid.ShouldBeEquivalentTo(originalYamlDoc.Uid);
        newYamlDoc.Title.ShouldBeEquivalentTo(originalYamlDoc.Title);
        newYamlDoc.Summary.ShouldBeEquivalentTo(originalYamlDoc.Summary);
        newYamlDoc.Abstract.ShouldBeEquivalentTo(originalYamlDoc.Abstract);
        newYamlDoc.Prerequisites.ShouldBeEquivalentTo(originalYamlDoc.Prerequisites);
        newYamlDoc.IconUrl.ShouldBeEquivalentTo(originalYamlDoc.IconUrl);
        newYamlDoc.Badge.ShouldBeEquivalentTo(originalYamlDoc.Badge);

        newYamlDoc.Levels.ShouldBeEquivalentTo(originalYamlDoc.Levels);
        newYamlDoc.Roles.ShouldBeEquivalentTo(originalYamlDoc.Roles);
        newYamlDoc.Products.ShouldBeEquivalentTo(originalYamlDoc.Products);
        newYamlDoc.Units.ShouldBeEquivalentTo(originalYamlDoc.Units);

        newYamlDoc.Metadata.Author.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Author);
        newYamlDoc.Metadata.Description.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Description);
        newYamlDoc.Metadata.MsAuthor.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsAuthor);
        //newYamlDoc.Metadata.MsDate.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsDate);
        newYamlDoc.Metadata.MsProduct.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsProduct);
        newYamlDoc.Metadata.MsTopic.ShouldBeEquivalentTo(originalYamlDoc.Metadata.MsTopic);
        newYamlDoc.Metadata.Title.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Title);
        //originalYamlDoc.Metadata.Robots.ShouldBeEquivalentTo(originalYamlDoc.Metadata.Robots);
    }
}


