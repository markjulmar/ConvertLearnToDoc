using ConvertLearnToDoc.Shared;
using ConvertLearnToDoc.Utility;
using Julmar.DocsToMarkdown;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using ContentRef = ConvertLearnToDoc.Utility.ContentRef;

#if USE_GITHUB_PAT
using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
#endif

namespace ConvertLearnToDoc.Controllers;

[Route("api/[Controller]")]
[ApiController]
[AuthorizeApi]
public class ContentConverterController(
#if USE_GITHUB_PAT
    ILogger<ContentConverterController> logger,
    IWebHostEnvironment hostingEnvironment,
    TokenService tokenService) : ControllerBase
#else
    ILogger<ContentConverterController> logger) : ControllerBase
#endif

{
    [HttpPost]
    [Route("toDoc")]
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

#if USE_GITHUB_PAT
        bool isLocal = hostingEnvironment.IsDevelopment();
        var gitHubToken = await tokenService.GetAccessTokenAsync() 
                            ?? Environment.GetEnvironmentVariable("GitHubToken");
        if (string.IsNullOrEmpty(gitHubToken) && !isLocal)
        {
            logger.LogError("ConvertLearnUrlToDocument: Cannot access GitHub organization or repo, are you a member of MicrosoftDocs?");
            return Unauthorized(new { message = "Unable to access GitHub API. Please contact support." });
        }

        var orgs = await GetOrganizationsAsync(gitHubToken);
        if (!orgs.Contains(contentRef.Organization))
        {
            logger.LogError("ConvertLearnUrlToDocument: User {User} is not a member of the organization {Organization}", user, contentRef.Organization);
            return Unauthorized(new { message = $"Cannot access GitHub organization or repo, are you a member of {contentRef.Organization}?" });
        }
#else
        var tempFolder = Path.Combine(Path.GetTempPath(), "LearnDocs");

        // Download the article or training module.
        logger.LogDebug("Downloading {InputFile}", request.Url);
        if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
        Directory.CreateDirectory(tempFolder);

        var downloader = new DocsConverter(tempFolder, new Uri(request.Url));
        var createdFiles = await downloader.ConvertAsync(logger: tag => logger.LogDebug("Skipped: {Tag}", tag.TrimStart().Substring(0,20)));
        if (createdFiles.Count == 0)
        {
            logger.LogError("ConvertLearnUrlToDocument: No files returned from ContentDownloader for {Url}", request.Url);
            return BadRequest(new { message = "Unable to identify content type at that URL on Microsoft Learn." });
        }
#endif

        var outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
        contentRef.Folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

        try
        {
            logger.LogDebug("ConvertLearnUrlToDocument({ContentRef}) => {OutputFile}", contentRef, outputFile);

            List<string>? errors;
            if (pageType == PageType.Article)
            {
#if USE_GITHUB_PAT

                errors = await SinglePageToDocx.ConvertFromRepoAsync(request.Url, contentRef.Organization,
                    contentRef.Repository, contentRef.Branch, contentRef.Folder, outputFile,
                    gitHubToken, new DocumentOptions { ZonePivot = request.ZonePivot??"" });
#else
                var markdownFile = createdFiles.Single(f => f.FileType == FileType.Markdown).Filename;
                errors = await SinglePageToDocx.ConvertFromFileAsync(request.Url, markdownFile, outputFile,
                    new DocumentOptions { ZonePivot = request.ZonePivot??"" });
#endif
            }
            else //if (pageType == PageType.Module)
            {
#if USE_GITHUB_PAT
                errors = await LearnToDocx.ConvertFromRepoAsync(request.Url,
                    contentRef.Organization, contentRef.Repository, contentRef.Branch,
                    contentRef.Folder, outputFile, gitHubToken,
                    new DocumentOptions {
                        ZonePivot = request.ZonePivot ?? "",
                        EmbedNotebookContent = request.EmbedNotebooks
                    });
#else
                var moduleFolder = createdFiles.First(f => f.FileType == FileType.Folder).Filename;
                errors = await LearnToDocx.ConvertFromFolderAsync(request.Url, moduleFolder, outputFile,
                    new DocumentOptions
                    {
                        ZonePivot = request.ZonePivot ?? "",
                        EmbedNotebookContent = request.EmbedNotebooks
                    });
#endif
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
        finally
        {
#if !USE_GITHUB_PAT
            logger.LogDebug("Cleaning up {TempFolder}", tempFolder);
            if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
#endif
        }
    }

    [HttpPost]
    [Route("toMarkdown")]
    public async Task<IActionResult> ConvertLearnUrlToMarkdown(LearnUrlConversionRequest request)
    {
        var user = ControllerExtensions.GetUsername(this.HttpContext);
        logger.LogInformation("ConvertLearnUrlToMarkdown({Request}) by {User}", request, user);

        if (!request.IsValid())
        {
            return BadRequest("Bad input - you must supply a URL to convert.");
        }


        var contentRef = await GetMetadataForLearnUrl(request.Url);
        if (!contentRef.IsValid())
        {
            return BadRequest("Bad input - the URL does not appear to be a valid article or training module.");
        }

        var tempFolder = Path.Combine(Path.GetTempPath(), "LearnDocs");

        try
        {
            // Download the article or training module.
            logger.LogDebug("Downloading {InputFile}", request.Url);
            if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
            Directory.CreateDirectory(tempFolder);

            var downloader = new DocsConverter(tempFolder, new Uri(request.Url));
            var createdFiles = await downloader.ConvertAsync(logger: tag => logger.LogDebug("Skipped: {Tag}", tag.TrimStart().Substring(0,20)));
            if (createdFiles.Count == 0)
            {
                logger.LogError("ConvertLearnUrlToMarkdown: No files returned from ContentDownloader for {Url}", request.Url);
                return BadRequest(new { message = "Unable to identify content type at that URL on Microsoft Learn." });
            }
            else
            {
                // Zip files.
                if (createdFiles.Count == 1)
                {
                    return await this.FileAttachment(createdFiles.Single().Filename, Constants.MarkdownMimeType);
                }
                else
                {
                    bool isArticle = createdFiles.Count(f => f.FileType == FileType.Markdown) == 1;
                    string baseFolder = isArticle
                        ? tempFolder
                        : createdFiles.First(f => f.FileType == FileType.Folder).Filename;
                    string baseFileName = isArticle
                        ? Path.GetFileNameWithoutExtension(createdFiles.Single(f => f.FileType == FileType.Markdown).Filename)
                        : Path.GetFileName(baseFolder);

                    var zipFile = Path.Combine(tempFolder, Path.ChangeExtension(baseFileName, "zip"));
                    if (System.IO.File.Exists(zipFile))
                    {
                        logger.LogDebug("DELETE {ZipFile} (existing)", zipFile);
                        System.IO.File.Delete(zipFile);
                    }

                    logger.LogDebug("ZIP {OutputPath} => {ZipFile}", baseFolder, zipFile);
                    System.IO.Compression.ZipFile.CreateFromDirectory(baseFolder, zipFile);

                    if (System.IO.File.Exists(zipFile))
                    {
                        logger.LogDebug("Returning .ZIP file");
                        return await this.FileAttachment(zipFile, Constants.ZipMimeType);
                    }
                }
            }
        }
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            ex.Handle(ex =>
            {
                logger.LogError("ConvertLearnUrlToMarkdown failed - {Exception}", ex);
                return true;
            });

            return BadRequest($"Unable to convert {contentRef} to a Markdown. Error: {ex.GetType()}: {ex.Message}");
        }
        catch (Exception ex)
        {
            logger.LogError("ConvertLearnUrlToMarkdown failed - {Exception}", ex);
            return BadRequest($"Unable to convert {contentRef} to Markdown. Error: {ex.GetType()}: {ex.Message}");
        }
        finally
        {
            logger.LogDebug("Cleaning up {TempFolder}", tempFolder);
            if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
        }

        return BadRequest(
            $"Unable to convert {request.Url} to Markdown - possibly an unsupported content type.");
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

#if USE_GITHUB_PAT
    private async Task<List<string>> GetOrganizationsAsync(string accessToken)
    {
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
#endif

}