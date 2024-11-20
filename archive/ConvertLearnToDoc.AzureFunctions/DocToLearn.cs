using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using ConvertLearnToDoc.AzureFunctions.Models;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace ConvertLearnToDoc.AzureFunctions;

public static class DocToLearn
{
    [FunctionName("DocToLearn")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        log.LogInformation("DocToLearn invoked.");

        var input = await req.ReadFormAsync();
        var model = new DocToLearnModel
        {
            WordDoc = input.Files[nameof(DocToLearnModel.WordDoc)],
            UseAsterisksForBullets = bool.TryParse(input[nameof(DocToLearnModel.UseAsterisksForBullets)], out var useAsterisksForBullets) && useAsterisksForBullets,
            UseAsterisksForEmphasis = bool.TryParse(input[nameof(DocToLearnModel.UseAsterisksForEmphasis)], out var useAsterisksForEmphasis) && useAsterisksForEmphasis,
            OrderedListUsesSequence = bool.TryParse(input[nameof(DocToLearnModel.OrderedListUsesSequence)], out var orderedListUsesSequence) && orderedListUsesSequence,
            UseIndentsForCodeBlocks = bool.TryParse(input[nameof(DocToLearnModel.UseIndentsForCodeBlocks)], out var useIndentsForCodeBlocks) && useIndentsForCodeBlocks,
            PrettyPipeTables = bool.TryParse(input[nameof(DocToLearnModel.PrettyPipeTables)], out var prettyPipeTables) && prettyPipeTables,
            IgnoreMetadata = bool.TryParse(input[nameof(DocToLearnModel.IgnoreMetadata)], out var ignoreMetadata) && ignoreMetadata,
            UseGenericIds = bool.TryParse(input[nameof(DocToLearnModel.UseGenericIds)], out var useGenericIds) && useGenericIds,
            UsePlainMarkdown = bool.TryParse(input[nameof(DocToLearnModel.UsePlainMarkdown)], out var usePlainMarkdown) && usePlainMarkdown,
        };

        var contentType = model.WordDoc?.ContentType;
        if (string.IsNullOrEmpty(model.WordDoc?.FileName) || contentType != Constants.WordMimeType)
        {
            return new BadRequestErrorMessageResult("Must pass in a valid .docx (Word) document.");
        }

        string baseFolder = Path.GetTempPath();
        string tempFile = Path.Combine(baseFolder, Path.GetTempFileName());

        // Copy the input file.
        await using (var stream = File.Create(tempFile))
        {
            await model.WordDoc.CopyToAsync(stream);
        }

        // Create the output folder.
        string moduleFolder = Path.GetFileNameWithoutExtension(model.WordDoc.FileName) ?? Constants.DefaultModuleName;
        if (Path.GetInvalidPathChars().Any(ch => moduleFolder.Contains(ch)))
        {
            moduleFolder = Constants.DefaultModuleName;
        }
        string outputPath = Path.Combine(baseFolder, moduleFolder);
        Directory.CreateDirectory(outputPath);

        try
        {
            log.LogDebug($"DocxToLearn(inputFile:{tempFile}, outputPath:{outputPath})");
            await DocxToLearn.ConvertAsync(tempFile, outputPath, new MarkdownOptions
            {
                UseAsterisksForBullets = model.UseAsterisksForBullets,
                UseAsterisksForEmphasis = model.UseAsterisksForEmphasis,
                OrderedListUsesSequence = model.OrderedListUsesSequence,
                UseIndentsForCodeBlocks = model.UseIndentsForCodeBlocks,
                IgnoreEmbeddedMetadata = model.IgnoreMetadata,
                UseGenericIds = model.UseGenericIds,
                UsePlainMarkdown = model.UsePlainMarkdown,
            });
        }
        catch (Exception ex)
        {
            log.LogError(ex.ToString());
            if (Directory.Exists(outputPath))
                Directory.Delete(outputPath, true);

            string errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return new BadRequestErrorMessageResult(
                $"Unable to convert {model.WordDoc.FileName}. {errorMessage}.");
        }

        string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
        if (File.Exists(zipFile))
        {
            log.LogDebug($"DELETE {zipFile}");
            File.Delete(zipFile);
        }
        log.LogDebug($"ZIP {outputPath} => {zipFile}");
        System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

        // Delete the temp stuff.
        log.LogDebug($"RMDIR {outputPath}");
        Directory.Delete(outputPath, true);
        log.LogDebug($"DELETE {tempFile}");
        File.Delete(tempFile);

        if (File.Exists(zipFile))
        {
            try
            {
                return new FileContentResult(await File.ReadAllBytesAsync(zipFile), "application/zip")
                    { FileDownloadName = Path.GetFileName(zipFile).Replace(' ', '-').ToLower() };
            }
            finally
            {
                File.Delete(zipFile);
            }

        }

        return new BadRequestErrorMessageResult(
            $"Unable to convert {model.WordDoc.FileName} to a Learn module.");
    }
}