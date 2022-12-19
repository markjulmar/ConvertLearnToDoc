using ConvertLearnToDoc.Shared;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;

namespace ConvertLearnToDoc.Server.Controllers;

[Route("api/[Controller]")]
[ApiController]
public class DocConverterController : ControllerBase
{
    private readonly ILogger<DocConverterController> logger;

    public DocConverterController(ILogger<DocConverterController> logger)
    {
        this.logger = logger;
    }

    [HttpPost]
    public async Task<IActionResult> Post([FromBody] ArticleOrModuleRef articleRef)
    {
        logger.LogInformation($"DocConverter: {articleRef}");

        if (articleRef.Document is not {ContentType: Constants.WordMimeType} 
            || string.IsNullOrWhiteSpace(articleRef.Document.FileName) 
            || articleRef.Document.Contents.Length <= 0)
        {
            return BadRequest("Bad input.");
        }

        try
        {
            if (articleRef.IsArticle)
            {
                return await ConvertArticleAsync(articleRef);
            }

            return await ConvertModuleAsync(articleRef);
        }
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            return BadRequest(
                $"Unable to convert {articleRef.Document.FileName}. {ex.GetType()}: {ex.Message}.");
        }
        catch (Exception ex)
        {
            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to convert {articleRef.Document.FileName}. {errorMessage}.");
        }
    }

    private async Task<IActionResult> ConvertModuleAsync(ArticleOrModuleRef articleRef)
    {
        string baseFolder = Path.GetTempPath();
        string tempFile = Path.Combine(baseFolder, Path.GetTempFileName());

        // Copy the input file.
        await using (var stream = System.IO.File.Create(tempFile))
        {
            await stream.WriteAsync(articleRef.Document!.Contents);
        }

        // Create the output folder.
        var moduleFolder = Path.GetFileNameWithoutExtension(articleRef.Document.FileName);
        if (Path.GetInvalidPathChars().Any(ch => moduleFolder.Contains(ch)))
        {
            moduleFolder = Constants.DefaultModuleName;
        }
        
        var outputPath = Path.Combine(baseFolder, moduleFolder);
        Directory.CreateDirectory(outputPath);

        try
        {
            logger.LogDebug($"DocxToLearn(inputFile:{tempFile}, outputPath:{outputPath})");
            await DocxToLearn.ConvertAsync(tempFile, outputPath, 
                new LearnMarkdownOptions {
                    UseAsterisksForBullets = articleRef.UseAsterisksForBullets,
                    UseAsterisksForEmphasis = articleRef.UseAsterisksForEmphasis,
                    OrderedListUsesSequence = articleRef.OrderedListUsesSequence,
                    UseIndentsForCodeBlocks = articleRef.UseIndentsForCodeBlocks,
                    IgnoreMetadata = articleRef.IgnoreMetadata,
                    UseGenericIds = articleRef.UseGenericIds,
                    UsePlainMarkdown = articleRef.UsePlainMarkdown,
                });
        }
        catch (Exception ex)
        {
            logger.LogError(ex.ToString());
            if (Directory.Exists(outputPath))
                Directory.Delete(outputPath, true);

            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to convert {articleRef.Document.FileName}. {errorMessage}.");
        }

        var zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
        if (System.IO.File.Exists(zipFile))
        {
            logger.LogDebug($"DELETE {zipFile}");
            System.IO.File.Delete(zipFile);
        }
        logger.LogDebug($"ZIP {outputPath} => {zipFile}");
        System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

        // Delete the temp stuff.
        logger.LogDebug($"RMDIR {outputPath}");
        Directory.Delete(outputPath, true);
        logger.LogDebug($"DELETE {tempFile}");
        System.IO.File.Delete(tempFile);

        if (System.IO.File.Exists(zipFile))
        {
            try
            {
                Response.Headers.Add("Content-Disposition", $"attachment;filename={Path.GetFileName(zipFile).Replace(' ', '-')}");
                return File(await System.IO.File.ReadAllBytesAsync(zipFile), Constants.ZipMimeType);
            }
            finally
            {
                System.IO.File.Delete(zipFile);
            }
        }

        return BadRequest(
            $"Unable to convert {articleRef.Document.FileName} to a training module.");
    }

    private async Task<IActionResult> ConvertArticleAsync(ArticleOrModuleRef articleRef)
    {
        var baseFolder = Path.GetTempPath();
        var tempFile = Path.Combine(baseFolder, Path.GetTempFileName());
        var baseFileName = Path.GetFileNameWithoutExtension(articleRef.Document!.FileName);
        if (Path.GetInvalidPathChars().Any(ch => baseFileName.Contains(ch)))
            baseFileName = Constants.DefaultModuleName;
        var outputPath = Path.Combine(baseFolder, baseFileName);

        try
        {
            // Copy the input file.
            await using (var stream = System.IO.File.Create(tempFile))
            {
                await stream.WriteAsync(articleRef.Document.Contents);
            }

            Directory.CreateDirectory(outputPath);

            // Get the markdown file.
            string markdownFile = Path.Combine(outputPath, Path.ChangeExtension(baseFileName, ".md"));

            // Convert the file to Markdown + media
            logger.LogDebug($"DocToPage(inputFile:{tempFile}, markdownFile:{markdownFile})");
            await DocxToSinglePage.ConvertAsync(tempFile, markdownFile, new MarkdownOptions
            {
                UseAsterisksForBullets = articleRef.UseAsterisksForBullets,
                UseAsterisksForEmphasis = articleRef.UseAsterisksForEmphasis,
                OrderedListUsesSequence = articleRef.OrderedListUsesSequence,
                UseIndentsForCodeBlocks = articleRef.UseIndentsForCodeBlocks
            }, articleRef.UsePlainMarkdown);

            // If we only produced a Markdown file, then return that.
            if (Directory.GetFiles(outputPath, "*.*", SearchOption.AllDirectories).Length == 1)
            {
                if (System.IO.File.Exists(markdownFile))
                {
                    logger.LogDebug($"Returning Markdown file.");

                    Response.Headers.Add("Content-Disposition", $"attachment;filename={Path.GetFileName(markdownFile).Replace(' ', '-')}");
                    return File(await System.IO.File.ReadAllBytesAsync(markdownFile), Constants.MarkdownMimeType);
                }
            }
            else
            {
                string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(baseFileName, "zip"));
                if (System.IO.File.Exists(zipFile))
                {
                    logger.LogDebug($"DELETE {zipFile} (existing).");
                    System.IO.File.Delete(zipFile);
                }

                logger.LogDebug($"ZIP {outputPath} => {zipFile}");
                System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

                if (System.IO.File.Exists(zipFile))
                {
                    logger.LogDebug($"Returning .ZIP file.");
                    Response.Headers.Add("Content-Disposition", $"attachment;filename={Path.GetFileName(zipFile).Replace(' ', '-')}");
                    return File(await System.IO.File.ReadAllBytesAsync(zipFile), Constants.ZipMimeType);
                }
            }
        }
        finally
        {
            // Delete the temp stuff.
            logger.LogDebug($"RMDIR {outputPath}");
            Directory.Delete(outputPath, true);
            logger.LogDebug($"DELETE {tempFile}");
            System.IO.File.Delete(tempFile);
        }

        return BadRequest(
            $"Unable to convert {articleRef.Document.FileName}. Possibly incorrect format?");
    }
}