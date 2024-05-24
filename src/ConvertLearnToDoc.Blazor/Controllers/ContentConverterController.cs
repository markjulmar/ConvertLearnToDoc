using ConvertLearnToDoc.Shared;
using ConvertLearnToDoc.Utility;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using ContentRef = ConvertLearnToDoc.Utility.ContentRef;

namespace ConvertLearnToDoc.Controllers;

[Route("api/[Controller]")]
[ApiController]
[AuthorizeApi]
public class ContentConverterController : ControllerBase
{
    private readonly ILogger<ContentConverterController> logger;
    private readonly IWebHostEnvironment hostingEnvironment;

    public ContentConverterController(ILogger<ContentConverterController> logger, IWebHostEnvironment hostingEnvironment)
    {
        this.logger = logger;
        this.hostingEnvironment = hostingEnvironment;
    }

    [HttpPost]
    public async Task<IActionResult> ConvertLearnUrlToDocument(LearnUrlConversionRequest request)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("ConvertLearnUrlToDocument({Request}) by {User}", request, user);

        if (!request.IsValid())
        {
            return BadRequest("Bad input - you must supply a URL to convert.");
        }

        var contentRef = await GetMetadataForLearnUrl(request.Url);

        if (!contentRef.IsValid())
        {
            return BadRequest("Bad input - the URL does not appear to be a valid article or training module.");
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
            logger.LogError("ConvertLearnUrlToDocument: Missing GitHubToken in server environment");
            return Unauthorized();
        }

        var outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
            contentRef.Folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

        try
        {
            logger.LogDebug("ConvertLearnUrlToDocument({ContentRef}) => {OutputFile}", contentRef, outputFile);

            List<string>? errors;
            if (pageType == PageType.Article)
            {
                errors = await SinglePageToDocx.ConvertFromRepoAsync(request.Url, contentRef.Organization,
                    contentRef.Repository, contentRef.Branch, contentRef.Folder, outputFile,
                    gitHubToken, new DocumentOptions { ZonePivot = request.ZonePivot??"" });
            }
            else //if (pageType == PageType.Module)
            {
                errors = await LearnToDocx.ConvertFromRepoAsync(request.Url,
                    contentRef.Organization, contentRef.Repository, contentRef.Branch,
                    contentRef.Folder, outputFile, gitHubToken,
                    new DocumentOptions {
                        ZonePivot = request.ZonePivot ?? "",
                        EmbedNotebookContent = request.EmbedNotebooks
                    });
            }
            if (System.IO.File.Exists(outputFile))
                return await this.FileAttachment(outputFile, Constants.WordMimeType);

            if (errors?.Count > 0)
                return BadRequest($"Unable to convert {contentRef} to a Word document.\r\n" +
                                  string.Join("\r\n", errors));

            return NotFound(
                $"Unable to convert {contentRef} to a Word document.");
        }
        catch (Exception ex)
        {
            logger.LogError("ConvertLearnUrlToDocument failed - {Exception}", ex);

            string errorMessage = ex.InnerException != null
                ? $"{ex.GetType()}: {ex.Message} ({ex.InnerException.GetType()}: {ex.InnerException.Message})"
                : $"{ex.GetType()}: {ex.Message}";

            return BadRequest($"Unable to convert {contentRef} to a Word document. Error: {errorMessage}");
        }
    }

    private async Task<ContentRef> GetMetadataForLearnUrl(string url)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("GetMetadataForLearnUrl({Url}) by {User}", url, user);

        if (string.IsNullOrEmpty(url)) 
            return new ContentRef();
        
        try
        {
            var metadata = await MSLearnRepos.DocsMetadata.LoadFromUrlAsync(url);
            string pageType = metadata.PageType??string.Empty;
            if (pageType.ToLower() == "learn")
                pageType += "." + metadata.PageKind;

            return new ContentRef
            {
                Organization = metadata.Organization ?? string.Empty,
                Repository = metadata.Repository ?? string.Empty,
                Branch = metadata.Branch ?? string.Empty,
                Folder = metadata.ContentPath ?? string.Empty,
                PageType = pageType,
            };
        }
        catch (Exception ex) 
        {
            logger.LogError("GetMetadata failed. {Exception}", ex);
        }

        return new ContentRef();
    }
}