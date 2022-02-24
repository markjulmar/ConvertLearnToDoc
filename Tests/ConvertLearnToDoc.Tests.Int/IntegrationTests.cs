using System.IO;
using Xunit;
using Shouldly;
using System.Threading.Tasks;
using System.Linq;
using FileComparisonLib;

namespace ConvertLearnToDoc.Tests.Int;

public class IntegrationTests
{
    [Fact]
    public async Task LearnToDoc_should_create_docx_file()
    {
        string learnModule = TestUtilities.CreateTestModuleFolder();
        string wordDoc = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetTempFileName(), "docx"));

        try
        {
            await TestUtilities.ConvertLearnModuleToWordDocumentAsync(learnModule, wordDoc);
            Assert.True(File.Exists(wordDoc));
            Assert.True((await File.ReadAllBytesAsync(wordDoc)).LongLength > 0);
        }
        finally
        {
            Directory.Delete(learnModule, true);
            File.Delete(wordDoc);
        }
    }

    [Fact]
    public async Task DocToLearn_should_create_yaml_markdown_files()
    {
        string wordDoc = TestUtilities.CreateTestWordDoc();
        string learnModule = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        string[] expectedFiles =
        {
            "1-introduction", 
            "2-application-deployment-types",
            "3-matrix-of-solutions", 
            "4-choosing-the-right-destination", 
            "6-summary"
        };

        try
        {
            await TestUtilities.ConvertWordDocumentToLearnFolderAsync(wordDoc, learnModule);

            var yamlFiles = Directory.GetFiles(learnModule, "*.yml", SearchOption.TopDirectoryOnly)
                                                .Select(Path.GetFileName)
                                                .ToList();

            Assert.NotEmpty(yamlFiles);
            foreach (var expected in expectedFiles.Select(fn => Path.ChangeExtension(fn, "yml")))
                yamlFiles.ShouldContain(expected);

            yamlFiles.ShouldContain("5-knowledge-check.yml");

            var mdFiles = Directory.GetFiles(Path.Combine(learnModule,"includes"), "*.md", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .ToList();

            Assert.NotEmpty(mdFiles);
            foreach (var expected in expectedFiles.Select(fn => Path.ChangeExtension(fn, "md")))
                mdFiles.ShouldContain(expected);

            mdFiles.ShouldNotContain("5-knowledge-check.md");
        }
        finally
        {
            Directory.Delete(learnModule, true);
            File.Delete(wordDoc);
        }
    }

    [Fact]
    public async Task Should_roundTrip()
    {
        string learnModule = TestUtilities.CreateTestModuleFolder();
        string wordDoc = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), "docx"));
        string generatedModule = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

        try
        {
            // Create the word doc.
            await TestUtilities.ConvertLearnModuleToWordDocumentAsync(learnModule, wordDoc);

            // Create the generated Learn module.
            await TestUtilities.ConvertWordDocumentToLearnFolderAsync(wordDoc, generatedModule);

            foreach (var originalFilePath in Directory.GetFiles(learnModule, "*.yml", SearchOption.TopDirectoryOnly))
            {
                var originalFileName = Path.GetFileName(originalFilePath);
                var generatedFilePath = Path.Combine(generatedModule, originalFileName);

                Assert.True(File.Exists(generatedFilePath));

                var comparison = FileComparer.Yaml(originalFilePath, generatedFilePath).ToList();
                comparison.ShouldBeEmpty();
            }

            // Do the same for Markdown content
            foreach (var originalFilePath in Directory.GetFiles(Path.Combine(learnModule, "includes"), "*.md", SearchOption.TopDirectoryOnly))
            {
                var originalFileName = Path.GetFileName(originalFilePath);
                var generatedFilePath = Path.Combine(generatedModule, "includes", originalFileName);

                Assert.True(File.Exists(generatedFilePath));

                var comparison = FileComparer.Markdown(originalFilePath, generatedFilePath).ToList();

                if (originalFileName == "2-application-deployment-types.md")
                {
                    Assert.Single(comparison);

                    // One diff on line 13 - an extra space which is removed.
                    var item = comparison.Single();
                    Assert.Equal(ChangeType.Changed, item.Change);
                    Assert.Equal("13/13", item.Key);
                    Assert.EndsWith("cloud provider. ", item.OriginalValue);
                    Assert.EndsWith("cloud provider.", item.NewValue);
                }
                else
                    Assert.Empty(comparison);

            }

        }
        finally
        {
            Directory.Delete(learnModule, true);
            Directory.Delete(generatedModule, true);
            File.Delete(wordDoc);
        }
    }
}


