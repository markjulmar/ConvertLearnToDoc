using System;
using System.IO;
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

        // Copy the input file.
        await using (var stream = File.Create(tempFile))
        {
            await model.WordDoc.CopyToAsync(stream);
        }

        // Create the output folder.
        string markdownFile = Path.ChangeExtension(
            Path.GetFileNameWithoutExtension(model.WordDoc.FileName) ?? Constants.DefaultModuleName, ".md");

        try
        {
            log.LogDebug($"DocToPage(inputFile:{tempFile}, markdownFile:{markdownFile})");
            await DocxToSinglePage.ConvertAsync(tempFile, markdownFile, new MarkdownOptions
            {
                UseAsterisksForBullets = model.UseAsterisksForBullets,
                UseAsterisksForEmphasis = model.UseAsterisksForEmphasis,
                OrderedListUsesSequence = model.OrderedListUsesSequence,
                UseIndentsForCodeBlocks = model.UseIndentsForCodeBlocks
            });
        }
        catch (Exception ex)
        {
            log.LogError(ex.ToString());
            return new BadRequestErrorMessageResult($"Error: {ex.Message}");
        }

        try
        {
            return new FileContentResult(await File.ReadAllBytesAsync(markdownFile), "text/markdown")
                { FileDownloadName = Path.GetFileName(markdownFile) };
        }
        finally
        {
            // Delete the temp stuff.
            log.LogDebug($"DELETE {markdownFile}");
            File.Delete(markdownFile);
            log.LogDebug($"DELETE {tempFile}");
            File.Delete(tempFile);
        }
    }
}