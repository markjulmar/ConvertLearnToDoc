using ConvertLearnToDoc.Shared;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;

namespace ConvertLearnToDoc.Server.Controllers;

[Route("api/[Controller]")]
[ApiController]
public class ContentConverterController : ControllerBase
{
    private readonly ILogger<ContentConverterController> logger;
    private readonly IWebHostEnvironment hostingEnvironment;

    public ContentConverterController(ILogger<ContentConverterController> logger, IWebHostEnvironment hostingEnvironment)
    {
        this.logger = logger;
        this.hostingEnvironment = hostingEnvironment;
    }

    [HttpGet]
    public async Task<ContentRef> Get(string url)
    {
        if (!string.IsNullOrEmpty(url))
        {
            try
            {
                var metadata = await MSLearnRepos.DocsMetadata.LoadFromUrlAsync(url);
                if (metadata != null)
                {
                    return new()
                    {
                        Organization = metadata.Organization ?? string.Empty,
                        Repository = metadata.Repository ?? string.Empty,
                        Branch = metadata.Branch ?? string.Empty,
                        Folder = metadata.ContentPath ?? string.Empty
                    };
                }
            }
            catch (Exception ex) 
            {
                logger.LogError(ex.ToString());
            }
        }

        return new();
    }

    [HttpPost]
    public async Task<IActionResult> Post([FromBody] ContentRef contentRef)
    {
        logger.LogInformation($"ContentConverter: {contentRef}");

        if (!contentRef.IsValid())
        {
            return BadRequest("Bad input.");
        }

        var pageType = PageType.Module;
        if (contentRef.Folder.EndsWith(Constants.MarkdownExtension, StringComparison.InvariantCultureIgnoreCase))
            pageType = PageType.Article;
        else
        {
            contentRef.Folder = Path.GetDirectoryName(contentRef.Folder) ?? contentRef.Folder;
        }

        contentRef.Folder = contentRef.Folder.Replace('\\', '/');
        if (!contentRef.Folder.StartsWith('/'))
            contentRef.Folder = '/' + contentRef.Folder;

        if (contentRef.Folder == "/")
        {
            return NotFound("Must identify an article or training module.");
        }

        bool isLocal = hostingEnvironment.IsDevelopment();
        var gitHubToken = Environment.GetEnvironmentVariable("GitHubToken");
        if (string.IsNullOrEmpty(gitHubToken) && !isLocal)
        {
            logger.LogError("Missing GitHubToken in server environment.");
            return Unauthorized();
        }

        var outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
            contentRef.Folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

        List<string>? errors = null;

        try
        {
            logger.LogDebug($"ConvertFromRepoAsync({contentRef}) => {outputFile}");

            if (pageType == PageType.Article)
            {
                errors = await SinglePageToDocx.ConvertFromRepoAsync(contentRef.Organization,
                    contentRef.Repository, contentRef.Branch, contentRef.Folder, outputFile,
                    gitHubToken, new DocumentOptions { ZonePivot = contentRef.ZonePivot });
            }
            else //if (pageType == PageType.Module)
            {
                errors = await LearnToDocx.ConvertFromRepoAsync(
                    contentRef.Organization, contentRef.Repository, contentRef.Branch,
                    contentRef.Folder, outputFile, gitHubToken,
                    new DocumentOptions {
                        ZonePivot = contentRef.ZonePivot ?? "",
                        EmbedNotebookContent = contentRef.EmbedNotebooks
                    });
            }
        }
        catch (Exception ex)
        {
            logger.LogError(ex.ToString());

            string errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest($"Unable to convert {contentRef.Organization}/{contentRef.Repository}:{contentRef.Branch}/{contentRef.Folder} to a Word document. Error: {errorMessage}");
        }

        if (System.IO.File.Exists(outputFile))
        {
            try
            {
                Response.Headers.Add("Content-Disposition", $"attachment;filename={Path.GetFileName(outputFile).Replace(' ', '-')}");
                return File(await System.IO.File.ReadAllBytesAsync(outputFile), Constants.WordMimeType);
            }
            finally
            {
                System.IO.File.Delete(outputFile);
            }
        }

        return NotFound(
            $"Unable to convert {contentRef.Organization}/{contentRef.Repository}:{contentRef.Branch}/{contentRef.Folder} to a Word document.");
    }
}