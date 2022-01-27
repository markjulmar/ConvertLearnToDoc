using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ConvertLearnToDocWeb.Models;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

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
            string outputFile = null;

            if (!string.IsNullOrEmpty(viewModel.ModuleUrl)
                && viewModel.ModuleUrl.ToLower().StartsWith("https"))
            {
                (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(viewModel.ModuleUrl);
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

            outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
                folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

            try
            {
                logger.LogDebug($"LearnToDocX(repo:{repo}, branch:{branch}, folder:{folder}: outputFile={outputFile})");
                await LearnToDocx.ConvertFromRepoAsync(repo, branch, folder, outputFile, viewModel.ZonePivot,
                    configuration.GetValue<string>("GitHub:Token"), new DocumentOptions { EmbedNotebookContent = viewModel.EmbedNotebookData });
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            if (System.IO.File.Exists(outputFile))
            {
                var fs = new FileStream(outputFile, FileMode.Open, FileAccess.Read);
                Response.RegisterForDispose(new TempFileRemover(logger, fs, outputFile));
                return File(fs, WordMimeType, Path.GetFileName(outputFile));
            }

            return View(nameof(Index));
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
            string baseFolder = Path.GetTempPath();
            string tempFile = Path.Combine(baseFolder, filename);

            // Copy the input file.
            await using (var stream = System.IO.File.Create(tempFile))
            {
                await viewModel.WordDoc.CopyToAsync(stream);
            }

            // Create the output folder.
            string moduleFolder = Path.GetFileNameWithoutExtension(filename);
            string outputPath = Path.Combine(baseFolder, moduleFolder);
            Directory.CreateDirectory(outputPath);

            try
            {
                logger.LogDebug($"DocxToLearn(inputFile:{tempFile}, outputPath:{outputPath})");
                await DocxToLearn.ConvertAsync(tempFile, outputPath, new MarkdownOptions {
                    UseAsterisksForBullets = viewModel.UseAsterisksForBullets,
                    UseAsterisksForEmphasis = viewModel.UseAsterisksForEmphasis,
                    OrderedListUsesSequence = viewModel.OrderedListUsesSequence,
                    UseAlternateHeaderSyntax = viewModel.UseAlternateHeaderSyntax,
                    UseIndentsForCodeBlocks = viewModel.UseIndentsForCodeBlocks
                });
            }
            catch (Exception ex)
            {
                logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
                return View(nameof(Index));
            }

            string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
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

            if (!System.IO.File.Exists(zipFile)) return View(nameof(Index));

            // Send back the zip file.
            var fs = new FileStream(zipFile, FileMode.Open, FileAccess.Read);
            Response.RegisterForDispose(new TempFileRemover(logger, fs, zipFile));
            return File(fs, "application/zip", Path.GetFileName(zipFile));
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error() => View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    internal class TempFileRemover : IDisposable
    {
        private readonly ILogger _logger;
        private readonly IDisposable _innerDispoable;
        private readonly string _path;

        public TempFileRemover(ILogger logger, IDisposable fs, string path)
        {
            _logger = logger;
            this._innerDispoable = fs;
            this._path = path;
        }

        public void Dispose()
        {
            _innerDispoable?.Dispose();
            if (File.Exists(_path))
            {
                try
                {
                    _logger.LogDebug($"DELETE {_path}");
                    File.Delete(_path);
                }
                catch
                {
                }
            }
        }
    }
}
