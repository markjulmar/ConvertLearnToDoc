using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ConvertLearnToDocWeb.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ConvertLearnToDocWeb.Controllers
{
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
        public async Task<IActionResult> ConvertLearnToDoc(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope($"ConvertLearnToDoc(Url: {viewModel.ModuleUrl}, Repo: {viewModel.GithubRepo}, Branch: {viewModel.GithubBranch}, Folder: \"{viewModel.GithubFolder}\")");

            if (!ModelState.IsValid)
            {
                return View(nameof(Index));
            }

            if (string.IsNullOrWhiteSpace(viewModel.ZonePivot))
                viewModel.ZonePivot = null;

            string repo, branch, folder;
            if (!string.IsNullOrEmpty(viewModel.ModuleUrl)
                && viewModel.ModuleUrl.ToLower().StartsWith("https"))
            {
                (repo, branch, folder) = await RetrieveLearnLocationFromUrlAsync(viewModel.ModuleUrl);
            }
            else if (!string.IsNullOrEmpty(viewModel.GithubRepo)
                && !string.IsNullOrEmpty(viewModel.GithubFolder))
            {
                repo = viewModel.GithubRepo;
                branch = string.IsNullOrEmpty(viewModel.GithubBranch) ? "live" : viewModel.GithubBranch;
                folder = viewModel.GithubFolder;
                if (!folder.StartsWith('/'))
                    folder = '/' + folder;
            }
            else
            {
                ModelState.AddModelError(string.Empty, "You must supply either a URL or GitHub specifics to identify a Learn module.");
                return View(nameof(Index));
            }

            try
            {
                logger.LogDebug($"LearnToDocX(repo:{repo}, branch:{branch}, folder:{folder})");
                var result = await CallLearnToDocConverter(new LearnToDocModel
                {
                    Repository = repo,
                    Branch = branch,
                    Folder = folder,
                    ZonePivot = viewModel.ZonePivot,
                    EmbedNotebookData = viewModel.EmbedNotebookData
                });

                if (result is {IsSuccessStatusCode: true})
                {
                    return new FileStreamResult(await result.Content.ReadAsStreamAsync(),
                        result.Content.Headers.ContentType?.MediaType ?? WordMimeType)
                    {
                        FileDownloadName = result.Content.Headers.ContentDisposition?.FileName ?? Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
                            folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"))

                };
                }

                throw new Exception(result.ReasonPhrase);
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            return View(nameof(Index));
        }

        private static async Task<(string repo, string branch, string folder)> RetrieveLearnLocationFromUrlAsync(string moduleUrl)
        {
            using var client = new HttpClient();
            string html = await client.GetStringAsync(moduleUrl);

            string pageKind = Regex.Match(html, @"<meta name=""page_kind"" content=""(.*?)""\s/>").Groups[1].Value;
            if (pageKind != "module")
                throw new ArgumentException("URL does not identify a Learn module - use the module landing page URL", nameof(moduleUrl));

            string lastCommit = Regex.Match(html, @"<meta name=""original_content_git_url"" content=""(.*?)""\s/>").Groups[1].Value;
            var uri = new Uri(lastCommit);
            if (uri.Host.ToLower() != "github.com")
                throw new ArgumentException("Identified module not hosted on GitHub", nameof(moduleUrl));

            var path = uri.LocalPath.ToLower().Split('/').Where(s => !string.IsNullOrWhiteSpace(s)).ToList();

            if (path[0] != "microsoftdocs")
                throw new ArgumentException("Identified module not in MicrosoftDocs organization", nameof(moduleUrl));

            string repo = path[1];
            if (!repo.StartsWith("learn-"))
                throw new ArgumentException("Identified module not in recognized MS Learn GitHub repo", nameof(moduleUrl));

            if (path.Last() == "index.yml")
                path.RemoveAt(path.Count - 1);

            string branch = path[3];
            string folder = string.Join('/', path.Skip(4));

            return (repo, branch, folder);
        }


        private async Task<HttpResponseMessage> CallLearnToDocConverter(LearnToDocModel model)
        {
            string endpoint = configuration.GetValue<string>("Service:LearnToDoc");
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException("Missing configuration for conversion function endpoints.");

            using var client = new HttpClient();
            return await client.PostAsync(endpoint, new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json"));
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocToLearn(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope("ConvertDocToLearn");

            var contentType = viewModel.WordDoc?.ContentType;
            if (string.IsNullOrEmpty(viewModel.WordDoc?.FileName) || contentType != WordMimeType)
            {
                ModelState.AddModelError(string.Empty, "You must upload a Word document.");
                return View(nameof(Index));
            }

            var filename = Path.GetFileName(viewModel.WordDoc.FileName);

            try
            {
                logger.LogDebug($"DocxToLearn(inputFile:{filename})");
                var result = await CallDocToLearnConverter(new DocToLearnModel
                {
                    WordDoc = viewModel.WordDoc,
                    OrderedListUsesSequence = viewModel.OrderedListUsesSequence,
                    UseAlternateHeaderSyntax = viewModel.UseAlternateHeaderSyntax,
                    UseAsterisksForBullets = viewModel.UseAsterisksForBullets,
                    UseAsterisksForEmphasis = viewModel.UseAsterisksForEmphasis,
                    UseIndentsForCodeBlocks = viewModel.UseIndentsForCodeBlocks
                });

                if (result is { IsSuccessStatusCode: true })
                {
                    return new FileStreamResult(await result.Content.ReadAsStreamAsync(),
                        result.Content.Headers.ContentType?.MediaType ?? "application/zip")
                    {
                        FileDownloadName = result.Content.Headers.ContentDisposition?.FileName 
                                           ?? Path.ChangeExtension(Path.GetFileNameWithoutExtension(filename), ".zip")
                    };
                }
                
                throw new Exception(result.ReasonPhrase);
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            return View(nameof(Index));
        }

        private async Task<HttpResponseMessage> CallDocToLearnConverter(DocToLearnModel model)
        {
            string endpoint = configuration.GetValue<string>("Service:DocToLearn");
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException("Missing configuration for conversion function endpoints.");

            var multiForm = new MultipartFormDataContent();
            multiForm.Add(new StringContent(model.UseAsterisksForBullets.ToString()), nameof(DocToLearnModel.UseAsterisksForBullets));
            multiForm.Add(new StringContent(model.UseAsterisksForEmphasis.ToString()), nameof(DocToLearnModel.UseAsterisksForEmphasis));
            multiForm.Add(new StringContent(model.OrderedListUsesSequence.ToString()), nameof(DocToLearnModel.OrderedListUsesSequence));
            multiForm.Add(new StringContent(model.UseAlternateHeaderSyntax.ToString()), nameof(DocToLearnModel.UseAlternateHeaderSyntax));
            multiForm.Add(new StringContent(model.UseIndentsForCodeBlocks.ToString()), nameof(DocToLearnModel.UseIndentsForCodeBlocks));
            
            var content = new StreamContent(model.WordDoc.OpenReadStream());
            content.Headers.ContentType = new MediaTypeHeaderValue(WordMimeType);
            multiForm.Add(content, nameof(DocToLearnModel.WordDoc), model.WordDoc.FileName);

            using var client = new HttpClient();
            return await client.PostAsync(endpoint, multiForm);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() => View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
