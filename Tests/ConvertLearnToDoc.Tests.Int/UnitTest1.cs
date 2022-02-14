using ConvertLearnToDoc.AzureFunctions;
using ConvertLearnToDoc.AzureFunctions.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text;
using System.Text.Json;
using Xunit;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Primitives;
using System.Collections.Generic;
using Microsoft.AspNetCore.Http.Internal;
using Shouldly;
using System.Threading.Tasks;
using LearnDocUtils;
using MSLearnRepos;
using System;
using System.IO.Compression;
using DiffPlex.DiffBuilder;
using System.Linq;
using System.Text.RegularExpressions;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using System.Diagnostics;

namespace ConvertLearnToDoc.Tests.Int;

public class UnitTest1
{
    private readonly ILogger logger = TestFactory.CreateLogger();

    [Fact]
    public async Task LearnToDoc_should_create_docx_file()
    {

        var doc = new LearnToDocModel()
        {
            Branch = "master",
            EmbedNotebookData = false,
            Folder = "/learn-pr/azure/route-and-process-data-logic-apps",
            Repository = "learn-pr",
            ZonePivot = ""
        };

        var fileContents = await CreateWordDoc(doc);
        fileContents.ShouldNotBeNull();
    }
    private async Task<byte[]> CreateWordDoc(LearnToDocModel model)
    {
        var request = TestFactory.CreateHttpRequest("", "");

        request.Method = "POST";
        var json = JsonSerializer.Serialize(model);
        byte[] bytes = Encoding.UTF8.GetBytes(json);
        MemoryStream stream = new(bytes);
        request.Body = stream;
        request.ContentLength = bytes.Length;
        request.ContentType = "application/x-www-form-urlencoded";

        var fileContentResult = (FileContentResult)await LearnToDoc.Run(request, logger);
        return fileContentResult.FileContents;
    }

    [Fact]
    public async Task DocToLearn_should_create_markdown_files()
    {
        var wordDocFilePath = $"D:\\dev\\ConvertLearnToDoc\\Tests\\ConvertLearnToDoc.Tests.Int\\route-and-process-data-logic-apps.docx";
        var fileName = "route-and-process-data-logic-apps.docx";

        var fileContents = await CreateZipFile(wordDocFilePath, fileName);
        fileContents.ShouldNotBeNull();
    }

    [Fact]
    public async Task Should_roundTrip()
    {
        var outputPath = Path.Combine(Path.GetTempPath(), "LearnToDocTests");

        //Generate and Download Word Doc from existing module
        var doc = new LearnToDocModel()
        {
            Branch = "master",
            EmbedNotebookData = false,
            Folder = "/learn-pr/azure/route-and-process-data-logic-apps",
            Repository = "learn-pr",
            ZonePivot = ""
        };

        var wordDocFileContents = await CreateWordDoc(doc);
        var wordDocTempFilePath = Path.Combine(outputPath, Path.GetTempFileName());
        wordDocTempFilePath = Path.ChangeExtension(wordDocTempFilePath, "docx");
        await File.WriteAllBytesAsync(wordDocTempFilePath, wordDocFileContents);

        //Generate markdown from Word doc
        var fileName = "route-and-process-data-logic-apps.docx";
        var zipFileContents = await CreateZipFile(wordDocTempFilePath, fileName);
        var zipFileTempFilePath = Path.Combine(outputPath, Path.GetTempFileName());
        zipFileTempFilePath = Path.ChangeExtension(zipFileTempFilePath, ".zip");
        await File.WriteAllBytesAsync(zipFileTempFilePath, zipFileContents);
        var unzippedPath = Path.Combine(outputPath, "LearnToDocTests", "ConvertedMD");
        if (Directory.Exists(unzippedPath))
        {
            Directory.Delete(unzippedPath, true);
        }
        ZipFile.ExtractToDirectory(zipFileTempFilePath, unzippedPath, overwriteFiles: true);

        //Download original source files
        var learn = new ModuleDownloader();
        var originalPath = Path.Combine(outputPath, "OriginalMD");

        var gitHubToken = Environment.GetEnvironmentVariable("GitHubToken");
        var tcService = TripleCrownGitHubService.CreateFromToken("learn-pr", "master", gitHubToken);

        var (module, markdownFile) = await learn
                                            .DownloadModuleAsync(tcService,
                                            learnFolder: "/learn-pr/azure/route-and-process-data-logic-apps",
                                            outputFolder: originalPath,
                                            embedNotebooks: false);


        var tempGHFolder = @"D:\dev\learn-pr\learn-pr\azure\route-and-process-data-logic-apps";

        foreach (var originalFilePath in Directory.GetFiles(tempGHFolder))
        {
            var originalFileName = Path.GetFileName(originalFilePath);
            if (File.Exists(Path.Combine(unzippedPath, originalFileName)))
            {
                Console.WriteLine(originalFileName);
                if (originalFileName.Contains(".yml", StringComparison.InvariantCultureIgnoreCase))
                {
                    var deserializer = new DeserializerBuilder()
                        .WithNamingConvention(CamelCaseNamingConvention.Instance)
                        .Build();


                    var originalText = File.ReadAllText(Path.Combine(tempGHFolder, originalFileName));
                    var originalTextReader = new StringReader(originalText);
                    var newText = File.ReadAllText(Path.Combine(unzippedPath, originalFileName));
                    var newTextReader = new StringReader(newText);

                    if (originalText.Contains("YamlMime:ModuleUnit", StringComparison.InvariantCultureIgnoreCase))
                    {
                        CompareModuleUnits(deserializer, originalTextReader, newTextReader, originalFileName);
                    }
                    else if (originalText.Contains("YamlMime:Module", StringComparison.InvariantCultureIgnoreCase))
                    {
                        CompareModuleIndexes(deserializer, originalTextReader, newTextReader, originalFileName);
                    }
                }

            }
        }

        //if (File.Exists(wordDocTempFilePath))
        //{
        //    File.Delete(wordDocTempFilePath);
        //}

        //if (File.Exists(zipFileTempFilePath))
        //{
        //    File.Delete(zipFileTempFilePath);
        //}
    }

    private void CompareModuleUnits(IDeserializer deserializer, StringReader originalTextReader, StringReader newTextReader, string fileName)
    {
        var originalYamlDoc = deserializer.Deserialize<YamlMimeModuleUnit.Root>(originalTextReader);
        originalYamlDoc.ShouldNotBeNull();
        originalYamlDoc.Metadata.ShouldNotBeNull();

        var newYamlDoc = deserializer.Deserialize<YamlMimeModuleUnit.Root>(newTextReader);
        newYamlDoc.ShouldNotBeNull();
        newYamlDoc.Metadata.ShouldNotBeNull();

        
        (originalYamlDoc.AzureSandbox == newYamlDoc.AzureSandbox || originalYamlDoc.Sandbox == newYamlDoc.Sandbox || originalYamlDoc.Sandbox == newYamlDoc.AzureSandbox || originalYamlDoc.AzureSandbox == newYamlDoc.Sandbox).ShouldBeTrue();

        //originalYamlDoc.AzureSandbox.ShouldBeEquivalentTo(newYamlDoc.AzureSandbox, $"AzureSandbox did not match. Expected {originalYamlDoc.AzureSandbox}, but was {newYamlDoc.AzureSandbox}. File {fileName}");
        //originalYamlDoc.Sandbox.ShouldBeEquivalentTo(newYamlDoc.Sandbox);

        originalYamlDoc.Content.ShouldBeEquivalentTo(newYamlDoc.Content);
        originalYamlDoc.DurationInMinutes.ShouldBeEquivalentTo(newYamlDoc.DurationInMinutes);
        originalYamlDoc.Title.ShouldBeEquivalentTo(newYamlDoc.Title);
        originalYamlDoc.Uid.ShouldBeEquivalentTo(newYamlDoc.Uid);

        originalYamlDoc.Metadata.Author.ShouldBeEquivalentTo(newYamlDoc.Metadata.Author);
        originalYamlDoc.Metadata.Description.ShouldBeEquivalentTo(newYamlDoc.Metadata.Description);
        originalYamlDoc.Metadata.MsAuthor.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsAuthor);
        //originalYamlDoc.Metadata.MsDate.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsDate);
        originalYamlDoc.Metadata.MsProd.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsProd);
        originalYamlDoc.Metadata.MsTopic.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsTopic);
        //originalYamlDoc.Metadata.Robots.ShouldBeEquivalentTo(newYamlDoc.Metadata.Robots);
        originalYamlDoc.Metadata.Title.ShouldBeEquivalentTo(newYamlDoc.Metadata.Title);
    }

    private void CompareModuleIndexes(IDeserializer deserializer, StringReader originalTextReader, StringReader newTextReader, string fileName)
    {
        var originalYamlDoc = deserializer.Deserialize<YamlMimeModule.Root>(originalTextReader);
        originalYamlDoc.ShouldNotBeNull();
        originalYamlDoc.Metadata.ShouldNotBeNull();

        var newYamlDoc = deserializer.Deserialize<YamlMimeModule.Root>(newTextReader);
        newYamlDoc.ShouldNotBeNull();
        newYamlDoc.Metadata.ShouldNotBeNull();


        originalYamlDoc.Uid.ShouldBeEquivalentTo(newYamlDoc.Uid);
        originalYamlDoc.Title.ShouldBeEquivalentTo(newYamlDoc.Title);
        originalYamlDoc.Summary.ShouldBeEquivalentTo(newYamlDoc.Summary);
        originalYamlDoc.Abstract.ShouldBeEquivalentTo(newYamlDoc.Abstract);
        originalYamlDoc.Prerequisites.ShouldBeEquivalentTo(newYamlDoc.Prerequisites);
        originalYamlDoc.IconUrl.ShouldBeEquivalentTo(newYamlDoc.IconUrl);
        originalYamlDoc.Badge.ShouldBeEquivalentTo(newYamlDoc.Badge);

        if (originalYamlDoc.Badge is not null && newYamlDoc.Badge is not null)
        {
            originalYamlDoc.Badge.Uid.ShouldBeEquivalentTo(newYamlDoc.Badge.Uid);
        }

        originalYamlDoc.Levels.ShouldBeEquivalentTo(newYamlDoc.Levels);
        originalYamlDoc.Roles.ShouldBeEquivalentTo(newYamlDoc.Roles);
        originalYamlDoc.Products.ShouldBeEquivalentTo(newYamlDoc.Products);
        originalYamlDoc.Units.ShouldBeEquivalentTo(newYamlDoc.Units);

        originalYamlDoc.Metadata.Author.ShouldBeEquivalentTo(newYamlDoc.Metadata.Author);
        originalYamlDoc.Metadata.Description.ShouldBeEquivalentTo(newYamlDoc.Metadata.Description);
        originalYamlDoc.Metadata.MsAuthor.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsAuthor);
        //originalYamlDoc.Metadata.MsDate.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsDate);
        originalYamlDoc.Metadata.MsProd.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsProd);
        originalYamlDoc.Metadata.MsTopic.ShouldBeEquivalentTo(newYamlDoc.Metadata.MsTopic);
        //originalYamlDoc.Metadata.Robots.ShouldBeEquivalentTo(newYamlDoc.Metadata.Robots);
        originalYamlDoc.Metadata.Title.ShouldBeEquivalentTo(newYamlDoc.Metadata.Title);
    }

    private async Task<byte[]> CreateZipFile(string wordDocFilePath, string fileName)
    {
        var request = TestFactory.CreateHttpRequest("", "");

        request.Method = "POST";

        var useAsterisksForBullets = true;
        var useAsterisksForEmphasis = true;
        var orderedListUsesSequence = true;
        var useAlternateHeaderSyntax = true;
        var useIndentsForCodeBlocks = true;

        request.ContentType = "application/x-www-form-urlencoded";

        var fileCollection = new FormFileCollection();
        var bytes = await File.ReadAllBytesAsync(wordDocFilePath);
        MemoryStream stream = new(bytes);

        var wordFile = new FormFile(stream, 0, bytes.Length, "wordDoc", fileName)
        {
            Headers = new HeaderDictionary
            {
                new KeyValuePair<string, StringValues>("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            }
        };
        fileCollection.Add(wordFile);

        var formCollection = new FormCollection(new Dictionary<string, StringValues>
            {
                        { "useAsterisksForBullets", useAsterisksForBullets.ToString() },
                        { "useAsterisksForEmphasis", useAsterisksForEmphasis.ToString() },
                        { "orderedListUsesSequence", orderedListUsesSequence.ToString() },
                        { "useAlternateHeaderSyntax", useAlternateHeaderSyntax.ToString() },
                        { "useIndentsForCodeBlocks", useIndentsForCodeBlocks.ToString() },
            }, fileCollection);

        request.Form = formCollection;

        var fileContentResult = (FileContentResult)await DocToLearn.Run(request, logger);

        return fileContentResult.FileContents;
    }
}


