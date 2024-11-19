using ConvertLearnToDoc.Shared;
using ConvertLearnToDoc.Utility;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using System.Net.Http.Headers;
using ContentRef = ConvertLearnToDoc.Utility.ContentRef;

namespace ConvertLearnToDoc.Controllers;

[Route("api/[Controller]")]
[ApiController]
[AuthorizeApi]
public class ContentConverterController : ControllerBase
{
    private readonly ILogger<ContentConverterController> logger;
    private readonly IWebHostEnvironment hostingEnvironment;
    private readonly TokenService tokenService;

    public ContentConverterController(
        ILogger<ContentConverterController> logger, 
        IWebHostEnvironment hostingEnvironment,
        TokenService tokenService)
    {
        this.logger = logger;
        this.hostingEnvironment = hostingEnvironment;
        this.tokenService = tokenService;
    }

    [HttpPost]
    public async Task<IActionResult> ConvertLearnUrlToDocument(LearnUrlConversionRequest request)
    {
        return BadRequest("Not Supported.");

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
        //var gitHubToken = await tokenService.GetAccessTokenAsync();
        if (string.IsNullOrEmpty(gitHubToken) && !isLocal)
        {
            logger.LogError("ConvertLearnUrlToDocument: Cannot access GitHub organization or repo, are you a member of MicrosoftDocs?");
            return Unauthorized(new { message = "Unable to access GitHub API. Please contact support." });
        }

        var orgs = await GetOrganizationsAsync();
        if (!orgs.Contains(contentRef.Organization))
        {
            logger.LogError("ConvertLearnUrlToDocument: User {User} is not a member of the organization {Organization}", user, contentRef.Organization);
            return Unauthorized(new { message = $"Cannot access GitHub organization or repo, are you a member of {contentRef.Organization}?" });
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
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            ex.Handle(ex =>
            {
                logger.LogError("ConvertLearnUrlToDocument failed - {Exception}", ex);
                return true;
            });

            return BadRequest($"Unable to convert {contentRef} to a Word document. Error: {ex.GetType()}: {ex.Message}");
        }
        catch (Exception ex)
        {
            logger.LogError("ConvertLearnUrlToDocument failed - {Exception}", ex);
            return BadRequest($"Unable to convert {contentRef} to a Word document. Error: {ex.GetType()}: {ex.Message}");
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

    private async Task<List<string>> GetOrganizationsAsync()
    {
        var accessToken = await tokenService.GetAccessTokenAsync();
        if (string.IsNullOrEmpty(accessToken))
            return [];

        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("YourAppName", "1.0"));

        var response = await client.GetAsync("https://api.github.com/user/orgs");
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        var organizations = System.Text.Json.JsonDocument.Parse(json)
            .RootElement
            .EnumerateArray()
            .Select(org => org.GetProperty("login").GetString() ?? "")
            .ToList();

        return organizations ?? [];
    }

}