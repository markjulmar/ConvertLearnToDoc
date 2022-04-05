using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ConvertLearnToDocWeb.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MSLearnRepos;
using Newtonsoft.Json;
using Activity = System.Diagnostics.Activity;

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
            bool isDocsPage = false;
            if (!string.IsNullOrEmpty(viewModel.ModuleUrl)
                && viewModel.ModuleUrl.ToLower().StartsWith("https"))
            {
                var md = await DocsMetadata.LoadFromUrlAsync(viewModel.ModuleUrl);
                repo = md.Repository;
                branch = md.Branch;
                folder = md.ContentPath;
                isDocsPage = md.PageType == "conceptual";
            }
            else if (!string.IsNullOrEmpty(viewModel.GithubRepo)
                && !string.IsNullOrEmpty(viewModel.GithubFolder))
            {
                repo = viewModel.GithubRepo;
                branch = string.IsNullOrEmpty(viewModel.GithubBranch) ? "live" : viewModel.GithubBranch;
                folder = viewModel.GithubFolder;
                if (!folder.StartsWith('/'))
                    folder = '/' + folder;
                isDocsPage = folder.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase);
            }
            else
            {
                ModelState.AddModelError(string.Empty, "You must supply either a URL or GitHub specifics to identify a Learn module or Docs conceptual page.");
                return View(nameof(Index));
            }

            try
            {
                var model = new LearnToDocModel
                {
                    Repository = repo,
                    Branch = branch,
                    Folder = isDocsPage ? folder : Path.GetDirectoryName(folder),
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

                throw new Exception(result.ReasonPhrase);
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            return View(nameof(Index));
        }

        private Task<HttpResponseMessage> CallPageToDocConverter(LearnToDocModel model) 
            => ToDocConverter(model, configuration.GetValue<string>("Service:PageToDoc"));
        private Task<HttpResponseMessage> CallLearnToDocConverter(LearnToDocModel model)
            => ToDocConverter(model, configuration.GetValue<string>("Service:LearnToDoc"));

        private async Task<HttpResponseMessage> ToDocConverter(LearnToDocModel model, string endpoint)
        {
            if (string.IsNullOrEmpty(endpoint))
                throw new InvalidOperationException("Missing configuration for conversion function endpoints.");
            using var client = new HttpClient();
            return await client.PostAsync(endpoint, new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json"));
        }

        private DocToLearnModel CreateToDocModel(ConversionViewModel viewModel)
        {
            var contentType = viewModel.WordDoc?.ContentType;
            if (string.IsNullOrEmpty(viewModel.WordDoc?.FileName) || contentType != WordMimeType)
                return null;

            return new DocToLearnModel
            {
                WordDoc = viewModel.WordDoc,
                OrderedListUsesSequence = viewModel.OrderedListUsesSequence,
                UseAsterisksForBullets = viewModel.UseAsterisksForBullets,
                UseAsterisksForEmphasis = viewModel.UseAsterisksForEmphasis,
                UseIndentsForCodeBlocks = viewModel.UseIndentsForCodeBlocks,
                PrettyPipeTables = viewModel.PrettyPipeTables
            };
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocToPage(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope("ConvertDocToPage");
            var model = CreateToDocModel(viewModel);
            if (model == null)
            {
                ModelState.AddModelError(string.Empty, "You must upload a Word document.");
                return View(nameof(Index));
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

                throw new Exception(result.ReasonPhrase);
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            return View(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocToLearn(ConversionViewModel viewModel)
        {
            using var scope = logger.BeginScope("ConvertDocToLearn");

            var model = CreateToDocModel(viewModel);
            if (model == null)
            {
                ModelState.AddModelError(string.Empty, "You must upload a Word document.");
                return View(nameof(Index));
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
                
                throw new Exception(result.ReasonPhrase);
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            return View(nameof(Index));
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
                { new StringContent(model.UseAsterisksForEmphasis.ToString()), nameof(DocToLearnModel.UseAsterisksForEmphasis) },
                { new StringContent(model.OrderedListUsesSequence.ToString()), nameof(DocToLearnModel.OrderedListUsesSequence) },
                { new StringContent(model.UseIndentsForCodeBlocks.ToString()), nameof(DocToLearnModel.UseIndentsForCodeBlocks) },
                { new StringContent(model.PrettyPipeTables.ToString()), nameof(DocToLearnModel.PrettyPipeTables) }
            };

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
