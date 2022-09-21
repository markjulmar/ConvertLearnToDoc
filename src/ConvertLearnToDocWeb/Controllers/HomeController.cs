using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ConvertLearnToDocWeb.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MSLearnRepos;
using Newtonsoft.Json;
using Activity = System.Diagnostics.Activity;

namespace ConvertLearnToDocWeb.Controllers
{
    public class BadResponse
    {
        public string Message { get; set; }
    }

    public class HomeController : Controller
    {
        const string WordMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

        private readonly ILogger<HomeController> logger;
        private readonly IConfiguration configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            this.logger = logger;
            this.configuration = configuration;
        }

        public IActionResult Index() => View();

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ConvertLearnToDoc(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope($"ConvertLearnToDoc(Url: {viewModel.ModuleUrl}, Org: {viewModel.GitHubOrg} + Repo: {viewModel.GithubRepo}, Branch: {viewModel.GithubBranch}, Folder: \"{viewModel.GithubFolder}\")");

            if (!string.IsNullOrEmpty(viewModel.TdRid))
                Response.Cookies.Append(viewModel.TdRid, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss:ff"), new CookieOptions { Expires = DateTimeOffset.Now + TimeSpan.FromMinutes(10) });

            viewModel.IsLearnToDoc = true;

            if (!ModelState.IsValid)
            {
                return View(nameof(Index), viewModel);
            }

            if (string.IsNullOrWhiteSpace(viewModel.ZonePivot))
                viewModel.ZonePivot = null;

            string org = null, repo = null, branch = null, folder = null;
            bool isDocsPage = false;
            if (!string.IsNullOrEmpty(viewModel.ModuleUrl)
                && viewModel.ModuleUrl.ToLower().StartsWith("https"))
            {
                DocsMetadata md = null;
                try
                {
                    md = await DocsMetadata.LoadFromUrlAsync(viewModel.ModuleUrl);
                }
                catch
                {
                    ViewData["ErrorMessage"] =
                        "You must supply either a public docs.microsoft.com URL or GitHub specifics to identify a Learn module or Docs conceptual page.";
                    return View(nameof(Index), viewModel);
                }

                org = md.Organization;
                repo = md.Repository;
                branch = md.Branch;
                folder = md.ContentPath;
                isDocsPage = md.PageType == "conceptual";
            }
            else if (!string.IsNullOrEmpty(viewModel.GithubRepo)
                && !string.IsNullOrEmpty(viewModel.GithubFolder))
            {
                org = string.IsNullOrEmpty(viewModel.GitHubOrg) ? Constants.DocsOrganization : viewModel.GitHubOrg;
                repo = viewModel.GithubRepo;
                branch = string.IsNullOrEmpty(viewModel.GithubBranch) ? "live" : viewModel.GithubBranch;
                folder = viewModel.GithubFolder;
                if (!folder.StartsWith('/'))
                    folder = '/' + folder;
                isDocsPage = folder.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase);
            }
            
            if (org == null || repo == null || branch == null || folder == null)
            {
                ViewData["ErrorMessage"] =
                    "You must supply either a public docs.microsoft.com URL or GitHub specifics to identify a Learn module or Docs conceptual page.";
                return View(nameof(Index), viewModel);
            }

            try
            {
                var model = new LearnToDocModel
                {
                    Organization = org,
                    Repository = repo,
                    Branch = branch,
                    Folder = isDocsPage ? folder : Path.HasExtension(folder) ? Path.GetDirectoryName(folder) : folder,
                    ZonePivot = viewModel.ZonePivot,
                    EmbedNotebookData = !isDocsPage && viewModel.EmbedNotebookData
                };

                HttpResponseMessage result;
                if (isDocsPage)
                {
                    logger.LogDebug($"SinglePageToDocX(repo:{repo}, branch:{branch}, folder:{folder})");
                    result = await CallPageToDocConverter(model);
                }
                else
                {
                    logger.LogDebug($"LearnToDocX(repo:{repo}, branch:{branch}, folder:{folder})");
                    result = await CallLearnToDocConverter(model);
                }

                if (result is {IsSuccessStatusCode: true})
                {
                    return new FileStreamResult(await result.Content.ReadAsStreamAsync(),
                        result.Content.Headers.ContentType?.MediaType ?? WordMimeType)
                    {
                        FileDownloadName = result.Content.Headers.ContentDisposition?.FileName ?? Path.Combine(Path.GetTempPath(), 
                            Path.ChangeExtension(folder.Split('/')
                                .Last(s => !string.IsNullOrWhiteSpace(s)), "docx"))
                    };
                }

                string message = await result.Content.ReadAsStringAsync();
                message = JsonConvert.DeserializeObject<BadResponse>(message)?.Message ?? message;
                ViewData["ErrorMessage"] = $"{result.ReasonPhrase}: {message}";
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ViewData["ErrorMessage"] = ex.Message;
            }

            return View(nameof(Index), viewModel);
        }

        private Task<HttpResponseMessage> CallPageToDocConverter(LearnToDocModel model) 
            => ToDocConverter(model, configuration.GetValue<string>("Service:PageToDoc"));
        private Task<HttpResponseMessage> CallLearnToDocConverter(LearnToDocModel model)
            => ToDocConverter(model, configuration.GetValue<string>("Service:LearnToDoc"));

        private static async Task<HttpResponseMessage> ToDocConverter(LearnToDocModel model, string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException("Missing configuration for conversion function endpoints.");
            using var client = new HttpClient();
            return await client.PostAsync(endpoint, new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json"));
        }

        private static DocToLearnModel CreateToDocModel(ConversionViewModel viewModel)
        {
            var contentType = viewModel.WordDoc?.ContentType;
            if (string.IsNullOrEmpty(viewModel.WordDoc?.FileName) || contentType != WordMimeType)
                return null;

            return new DocToLearnModel
            {
                WordDoc = viewModel.WordDoc,
                UsePlainMarkdown = viewModel.UsePlainMarkdown,
                OrderedListUsesSequence = viewModel.OrderedListUsesSequence,
                UseAsterisksForBullets = viewModel.UseAsterisksForBullets,
                UseAsterisksForEmphasis = viewModel.UseAsterisksForEmphasis,
                UseIndentsForCodeBlocks = viewModel.UseIndentsForCodeBlocks,
                PrettyPipeTables = viewModel.PrettyPipeTables
            };
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ConvertDocToPage(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope("ConvertDocToPage");

            if (!string.IsNullOrEmpty(viewModel.FdRid))
                Response.Cookies.Append(viewModel.FdRid, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss:ff"), new CookieOptions { Expires = DateTimeOffset.Now + TimeSpan.FromMinutes(10) });
            viewModel.IsLearnToDoc = false;

            var model = CreateToDocModel(viewModel);
            if (model == null)
            {
                ViewData["ErrorMessage"] = "You must upload a Word document.";
                return View(nameof(Index), viewModel);
            }

            var filename = Path.GetFileName(viewModel.WordDoc.FileName);

            try
            {
                logger.LogDebug($"ConvertDocToPage(inputFile:{filename})");
                var result = await CallDocToPageConverter(model);

                if (result is { IsSuccessStatusCode: true })
                {
                    return new FileStreamResult(await result.Content.ReadAsStreamAsync(),
                        result.Content.Headers.ContentType?.MediaType ?? "text/markdown")
                    {
                        FileDownloadName = result.Content.Headers.ContentDisposition?.FileName
                                           ?? Path.ChangeExtension(Path.GetFileNameWithoutExtension(filename), ".md")
                    };
                }

                string message = await result.Content.ReadAsStringAsync();
                message = JsonConvert.DeserializeObject<BadResponse>(message)?.Message ?? message;
                ViewData["ErrorMessage"] = $"{result.ReasonPhrase}: {message}";
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ViewData["ErrorMessage"] = ex.Message;
            }

            return View(nameof(Index), viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ConvertDocToLearn(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope("ConvertDocToLearn");

            if (!string.IsNullOrEmpty(viewModel.FdRid))
                Response.Cookies.Append(viewModel.FdRid, DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss:ff"), new CookieOptions { Expires = DateTimeOffset.Now + TimeSpan.FromMinutes(10) });
            viewModel.IsLearnToDoc = false;

            var model = CreateToDocModel(viewModel);
            if (model == null)
            {
                ViewData["ErrorMessage"] = "You must upload a Word document.";
                return View(nameof(Index), viewModel);
            }

            var filename = Path.GetFileName(viewModel.WordDoc.FileName);

            try
            {
                logger.LogDebug($"DocxToLearn(inputFile:{filename})");
                var result = await CallDocToLearnConverter(model);

                if (result is { IsSuccessStatusCode: true })
                {
                    return new FileStreamResult(await result.Content.ReadAsStreamAsync(),
                        result.Content.Headers.ContentType?.MediaType ?? "application/zip")
                    {
                        FileDownloadName = result.Content.Headers.ContentDisposition?.FileName 
                                           ?? Path.ChangeExtension(Path.GetFileNameWithoutExtension(filename), ".zip")
                    };
                }

                string message = await result.Content.ReadAsStringAsync();
                message = JsonConvert.DeserializeObject<BadResponse>(message)?.Message ?? message;
                ViewData["ErrorMessage"] = $"{result.ReasonPhrase}: {message}";
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ViewData["ErrorMessage"] = ex.Message;
            }

            return View(nameof(Index), viewModel);
        }

        private Task<HttpResponseMessage> CallDocToLearnConverter(DocToLearnModel model) 
            => CallDocConverter(model, configuration.GetValue<string>("Service:DocToLearn"));

        private Task<HttpResponseMessage> CallDocToPageConverter(DocToLearnModel model)
            => CallDocConverter(model, configuration.GetValue<string>("Service:DocToPage"));

        private static async Task<HttpResponseMessage> CallDocConverter(DocToLearnModel model, string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException("Missing configuration for conversion function endpoints.");

            var multiForm = new MultipartFormDataContent
            {
                { new StringContent(model.UseAsterisksForBullets.ToString()), nameof(DocToLearnModel.UseAsterisksForBullets) },
                { new StringContent(model.UsePlainMarkdown.ToString()), nameof(DocToLearnModel.UsePlainMarkdown) },
                { new StringContent(model.UseAsterisksForEmphasis.ToString()), nameof(DocToLearnModel.UseAsterisksForEmphasis) },
                { new StringContent(model.OrderedListUsesSequence.ToString()), nameof(DocToLearnModel.OrderedListUsesSequence) },
                { new StringContent(model.UseIndentsForCodeBlocks.ToString()), nameof(DocToLearnModel.UseIndentsForCodeBlocks) },
                { new StringContent(model.PrettyPipeTables.ToString()), nameof(DocToLearnModel.PrettyPipeTables) }
            };

            var content = new StreamContent(model.WordDoc.OpenReadStream());
            content.Headers.ContentType = new MediaTypeHeaderValue(WordMimeType);
            multiForm.Add(content, nameof(DocToLearnModel.WordDoc), model.WordDoc.FileName);

            using var client = new HttpClient {Timeout = TimeSpan.FromMinutes(5)};
            return await client.PostAsync(endpoint, multiForm);
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() => View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
