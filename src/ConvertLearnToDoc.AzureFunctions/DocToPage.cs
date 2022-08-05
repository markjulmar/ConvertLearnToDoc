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

public static class DocToPage
{
    [FunctionName("DocToPage")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        log.LogInformation("DocToPage invoked.");

        var input = await req.ReadFormAsync();
        var model = new DocToLearnModel
        {
            WordDoc = input.Files[nameof(DocToLearnModel.WordDoc)],
            UsePlainMarkdown = bool.TryParse(input[nameof(DocToLearnModel.UsePlainMarkdown)], out var usePlainMarkdown) && usePlainMarkdown,
            UseAsterisksForBullets = bool.TryParse(input[nameof(DocToLearnModel.UseAsterisksForBullets)], out var useAsterisksForBullets) && useAsterisksForBullets,
            UseAsterisksForEmphasis = bool.TryParse(input[nameof(DocToLearnModel.UseAsterisksForEmphasis)], out var useAsterisksForEmphasis) && useAsterisksForEmphasis,
            OrderedListUsesSequence = bool.TryParse(input[nameof(DocToLearnModel.OrderedListUsesSequence)], out var orderedListUsesSequence) && orderedListUsesSequence,
            UseIndentsForCodeBlocks = bool.TryParse(input[nameof(DocToLearnModel.UseIndentsForCodeBlocks)], out var useIndentsForCodeBlocks) && useIndentsForCodeBlocks,
            PrettyPipeTables = bool.TryParse(input[nameof(DocToLearnModel.PrettyPipeTables)], out var prettyPipeTables) && prettyPipeTables
        };

        var contentType = model.WordDoc?.ContentType;
        if (string.IsNullOrEmpty(model.WordDoc?.FileName) || contentType != Constants.WordMimeType)
        {
            return new BadRequestErrorMessageResult("Invalid request.");
        }

        string baseFolder = Path.GetTempPath();
        string tempFile = Path.Combine(baseFolder, Path.GetTempFileName());
        string baseFileName = Path.GetFileNameWithoutExtension(model.WordDoc.FileName) ?? Constants.DefaultModuleName;
        if (Path.GetInvalidPathChars().Any(ch => baseFileName.Contains(ch)))
            baseFileName = Constants.DefaultModuleName;
        string outputPath = Path.Combine(baseFolder, baseFileName);

        try
        {
            // Copy the input file.
            await using (var stream = File.Create(tempFile))
            {
                await model.WordDoc.CopyToAsync(stream);
            }

            Directory.CreateDirectory(outputPath);

            // Get the markdown file.
            string markdownFile = Path.Combine(outputPath, Path.ChangeExtension(baseFileName, ".md"));

            // Convert the file to Markdown + media
            log.LogDebug($"DocToPage(inputFile:{tempFile}, markdownFile:{markdownFile})");
            await DocxToSinglePage.ConvertAsync(tempFile, markdownFile, new MarkdownOptions
            {
                UseAsterisksForBullets = model.UseAsterisksForBullets,
                UseAsterisksForEmphasis = model.UseAsterisksForEmphasis,
                OrderedListUsesSequence = model.OrderedListUsesSequence,
                UseIndentsForCodeBlocks = model.UseIndentsForCodeBlocks
            }, model.UsePlainMarkdown);

            // If we only produced a Markdown file, then return that.
            if (Directory.GetFiles(outputPath, "*.*", SearchOption.AllDirectories).Length == 1)
            {
                if (File.Exists(markdownFile))
                {
                    log.LogDebug($"Returning Markdown file.");
                    return new FileContentResult(await File.ReadAllBytesAsync(markdownFile), "text/markdown")
                        { FileDownloadName = Path.GetFileName(markdownFile) };
                }
            }
            else
            {
                string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(baseFileName, "zip"));
                if (File.Exists(zipFile))
                {
                    log.LogDebug($"DELETE {zipFile} (existing).");
                    File.Delete(zipFile);
                }

                log.LogDebug($"ZIP {outputPath} => {zipFile}");
                System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

                if (File.Exists(zipFile))
                {
                    log.LogDebug($"Returning .ZIP file.");
                    return new FileContentResult(await File.ReadAllBytesAsync(zipFile), "application/zip")
                        {FileDownloadName = Path.GetFileName(zipFile).Replace(' ','-').ToLower()};
                }
            }
        }
        catch (AggregateException aex)
        {
            return new BadRequestErrorMessageResult(
                $"Unable to convert {model.WordDoc.FileName}. {aex.Flatten().Message}.");
        }
        catch (Exception ex)
        {
            return new BadRequestErrorMessageResult(
                $"Unable to convert {model.WordDoc.FileName}. {ex.Message}.");
        }
        finally
        {
            // Delete the temp stuff.
            log.LogDebug($"RMDIR {outputPath}");
            Directory.Delete(outputPath, true);
            log.LogDebug($"DELETE {tempFile}");
            File.Delete(tempFile);
        }

        return new BadRequestErrorMessageResult(
            $"Unable to convert {model.WordDoc.FileName}. Possibly incorrect format?");
    }
}