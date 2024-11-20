using ConvertLearnToDoc.Shared;
using ConvertLearnToDoc.Utility;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;

namespace ConvertLearnToDoc.Controllers;

[Route("api/[Controller]")]
[ApiController]
[AuthorizeApi]
public class DocConverterController : ControllerBase
{
    private readonly ILogger<DocConverterController> logger;

    public DocConverterController(ILogger<DocConverterController> logger)
    {
        this.logger = logger;
    }

    private static bool IsValidDocument(BrowserFile? document)
    {
        return document is {ContentType: Constants.WordMimeType} 
               && !string.IsNullOrWhiteSpace(document.FileName) 
               && document.Contents.Length > 100;
    }

    [HttpPost]
    [Route("metadata")]
    public async Task<IActionResult> GetModuleInfoFromDocumentAsync([FromBody] BrowserFile document)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("GetModuleInfoFromDocumentAsync({FileName}) by {User}", document.FileName, user);

        if (!IsValidDocument(document))
            return BadRequest("Document is not valid");

        var baseFolder = Path.GetTempPath();
        var tempFile = Path.Combine(baseFolder, Path.GetTempFileName());

        // Copy the input file.
        await using (var stream = System.IO.File.Create(tempFile))
        {
            await stream.WriteAsync(document.Contents);
        }

        try
        {
            if (document.IsArticle)
            {
                var dictionary = DocsDownloader.GetArticleMetadataFromDocument(tempFile);
                if (dictionary.Count > 0)
                {
                    return Ok(PersistenceUtilities.ObjectToYamlString(dictionary));
                }
            }
            else
            {
                var md = ModuleBuilder.LoadDocumentMetadata(tempFile, false, false, null);
                if (md.ModuleData != null)
                {
                    return Ok(PersistenceUtilities.ObjectToYamlString(md.ModuleData));
                }
            }

            return NotFound();
        }
        catch (FileFormatException)
        {
            return BadRequest(
                $"Unable to read metadata from {document.FileName}, Invalid format or document is protected.");
        }
        catch (Exception ex)
        {
            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to read metadata from {document.FileName}. {errorMessage}.");
        }
        finally
        {
            System.IO.File.Delete(tempFile);
        }
    }


    [HttpPost]
    [Route("article")]
    public async Task<IActionResult> ConvertArticleFromDocumentAsync([FromBody] ArticleRef article)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("ConvertArticleFromDocumentAsync({FileName}) by {User}", article.Document?.FileName ?? "<unknown>", user);

        if (!IsValidDocument(article.Document))
            return BadRequest("Document is not valid");

        try
        {
            return await ConvertArticleAsync(article);
        }
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            return BadRequest(
                $"Unable to convert {article.Document?.FileName}. {ex.GetType()}: {ex.Message}.");
        }
        catch (FileFormatException)
        {
            return BadRequest(
                $"Unable to convert {article.Document?.FileName}, Invalid format or document is protected.");
        }
        catch (Exception ex)
        {
            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to convert {article.Document?.FileName}. {errorMessage}.");
        }
    }

    [HttpPost]
    [Route("module")]
    public async Task<IActionResult> ConvertModuleFromDocumentAsync([FromBody] ModuleRef module)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("ConvertModuleFromDocumentAsync({ModuleRef}) by {User}", module, user);

        if (!IsValidDocument(module.Document))
            return BadRequest("Document is not valid");

        try
        {
            return await ConvertModuleAsync(module);
        }
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            return BadRequest(
                $"Unable to convert {module.Document?.FileName}. {ex.GetType()}: {ex.Message}.");
        }
        catch (FileFormatException)
        {
            return BadRequest(
                $"Unable to convert {module.Document?.FileName}, Invalid format or document is protected.");
        }
        catch (Exception ex)
        {
            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to convert {module.Document?.FileName}. {errorMessage}.");
        }
    }

    private async Task<IActionResult> ConvertModuleAsync(ModuleRef moduleRef)
    {
        var baseFolder = Path.GetTempPath();
        var tempFile = Path.Combine(baseFolder, Path.GetTempFileName());

        // Copy the input file.
        await using (var stream = System.IO.File.Create(tempFile))
        {
            await stream.WriteAsync(moduleRef.Document!.Contents);
        }

        // Create the output folder.
        var moduleFolder = Path.GetFileNameWithoutExtension(moduleRef.Document.FileName);
        if (Path.GetInvalidPathChars().Any(ch => moduleFolder.Contains(ch)))
        {
            moduleFolder = Constants.DefaultModuleName;
        }
        
        var outputPath = Path.Combine(baseFolder, moduleFolder);
        Directory.CreateDirectory(outputPath);

        try
        {
            logger.LogDebug("DocxToLearn({InputFile}) => {OutputFile})", tempFile, outputPath);
            await DocxToLearn.ConvertAsync(tempFile, outputPath, 
                new LearnDocUtils.MarkdownOptions {
                    UseAsterisksForBullets = moduleRef.UseAsterisksForBullets,
                    UseAsterisksForEmphasis = moduleRef.UseAsterisksForEmphasis,
                    OrderedListUsesSequence = moduleRef.OrderedListUsesSequence,
                    UseIndentsForCodeBlocks = moduleRef.UseIndentsForCodeBlocks,
                    IgnoreEmbeddedMetadata = moduleRef.IgnoreMetadata,
                    Metadata = moduleRef.Metadata,
                    UseGenericIds = moduleRef.UseGenericIds,
                    UsePlainMarkdown = moduleRef.UsePlainMarkdown,
                });
        }
        catch (Exception ex)
        {
            logger.LogError("DocxToLearn failed, {Exception}", ex);
            if (Directory.Exists(outputPath))
                Directory.Delete(outputPath, true);

            var errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest(
                $"Unable to convert {moduleRef.Document.FileName}. {errorMessage}.");
        }

        var zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
        if (System.IO.File.Exists(zipFile))
        {
            logger.LogDebug("DELETE {ZipFile}", zipFile);
            System.IO.File.Delete(zipFile);
        }
        logger.LogDebug("ZIP {OutputPath} => {ZipFile}", outputPath, zipFile);
        System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

        // Delete the temp stuff.
        logger.LogDebug("RMDIR {OutputPath}", outputPath);
        Directory.Delete(outputPath, true);
        logger.LogDebug("DELETE {InputFile}", tempFile);
        System.IO.File.Delete(tempFile);

        if (System.IO.File.Exists(zipFile))
            return await this.FileAttachment(zipFile, Constants.ZipMimeType);

        return BadRequest(
            $"Unable to convert {moduleRef.Document.FileName} to a training module.");
    }

    private async Task<IActionResult> ConvertArticleAsync(ArticleRef articleRef)
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
            var markdownFile = Path.Combine(outputPath, Path.ChangeExtension(baseFileName, ".md"));

            // Convert the file to Markdown + media
            logger.LogDebug("DocToPage({InputFile}) => {MarkdownFile}", tempFile, markdownFile);
            await DocxToSinglePage.ConvertAsync(tempFile, markdownFile, 
                new LearnDocUtils.MarkdownOptions
            {
                UseAsterisksForBullets = articleRef.UseAsterisksForBullets,
                UseAsterisksForEmphasis = articleRef.UseAsterisksForEmphasis,
                OrderedListUsesSequence = articleRef.OrderedListUsesSequence,
                UseIndentsForCodeBlocks = articleRef.UseIndentsForCodeBlocks,
                UsePlainMarkdown = articleRef.UsePlainMarkdown,
                Metadata = articleRef.Metadata,
            });

            // If we only produced a Markdown file, then return that.
            if (Directory.GetFiles(outputPath, "*.*", SearchOption.AllDirectories).Length == 1)
            {
                if (System.IO.File.Exists(markdownFile))
                {
                    logger.LogDebug("Returning Markdown file");
                    return await this.FileAttachment(markdownFile, Constants.MarkdownMimeType);
                }
            }
            else
            {
                var zipFile = Path.Combine(baseFolder, Path.ChangeExtension(baseFileName, "zip"));
                if (System.IO.File.Exists(zipFile))
                {
                    logger.LogDebug("DELETE {ZipFile} (existing)", zipFile);
                    System.IO.File.Delete(zipFile);
                }

                logger.LogDebug("ZIP {OutputPath} => {ZipFile}", outputPath, zipFile);
                System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

                if (System.IO.File.Exists(zipFile))
                {
                    logger.LogDebug("Returning .ZIP file");
                    return await this.FileAttachment(zipFile, Constants.ZipMimeType);
                }
            }
        }
        finally
        {
            // Delete the temp stuff.
            logger.LogDebug("RMDIR {OutputPath}", outputPath);
            Directory.Delete(outputPath, true);
            logger.LogDebug("DELETE {TempFile}", tempFile);
            System.IO.File.Delete(tempFile);
        }

        return BadRequest(
            $"Unable to convert {articleRef.Document.FileName}. Possibly incorrect format?");
    }
}