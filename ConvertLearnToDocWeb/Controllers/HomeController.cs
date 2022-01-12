using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ConvertLearnToDocWeb.Models;
using LearnDocUtils;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace ConvertLearnToDocWeb.Controllers
{
    public class HomeController : Controller
    {
        const string WordMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index() => View();

        [HttpPost]
        public async Task<IActionResult> ConvertLearnToDoc(LearnToDocViewModel model)
        {
            using var scope = _logger.BeginScope($"ConvertLearnToDoc(Url: {model.ModuleUrl}, Repo: {model.GithubRepo}, Branch: {model.GithubBranch}, Folder: \"{model.GithubFolder}\")");

            if (!ModelState.IsValid)
            {
                return View(nameof(Index));
            }

            string repo, branch, folder;
            string outputFile = null;

            if (!string.IsNullOrEmpty(model.ModuleUrl)
                && model.ModuleUrl.ToLower().StartsWith("https"))
            {
                (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(model.ModuleUrl);
            }
            else if (!string.IsNullOrEmpty(model.GithubRepo)
                && !string.IsNullOrEmpty(model.GithubFolder))
            {
                repo = model.GithubRepo;
                branch = string.IsNullOrEmpty(model.GithubBranch) ? "live" : model.GithubBranch;
                folder = model.GithubFolder;
                if (!folder.StartsWith('/'))
                    folder = '/' + folder;
            }
            else
            {
                ModelState.AddModelError(string.Empty, "You must supply either a URL or GitHub specifics to identify a Learn module.");
                return View(nameof(Index));
            }

            outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

            try
            {
                _logger.LogDebug($"LearnToDocX(repo:{repo}, branch:{branch}, folder:{folder}: outputFile={outputFile})");
                await LearnToDocx.ConvertFromRepoAsync(repo, branch, folder, outputFile, null,
                    _configuration.GetValue<string>("GitHub:Token"), false,
                    model.UseLegacyConverter
                    ? MarkdownConverterFactory.WithPandoc
                    : MarkdownConverterFactory.WithDxPlus);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            if (System.IO.File.Exists(outputFile))
            {
                var fs = new FileStream(outputFile, FileMode.Open, FileAccess.Read);
                Response.RegisterForDispose(new TempFileRemover(_logger, fs, outputFile));
                return File(fs, WordMimeType, Path.GetFileName(outputFile));
            }

            return View(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocToLearn(IFormFile wordDoc, bool useLegacyConverter)
        {
            using var scope = _logger.BeginScope("ConvertDocToLearn");

            var contentType = wordDoc?.ContentType;

            if (string.IsNullOrEmpty(wordDoc?.FileName) || contentType != WordMimeType)
            {
                ModelState.AddModelError(string.Empty, "You must upload a Word document.");
                return View(nameof(Index));
            }

            var filename = Path.GetFileName(wordDoc.FileName);
            string baseFolder = Path.GetTempPath();
            string tempFile = Path.Combine(baseFolder, filename);

            // Copy the input file.
            await using (var stream = System.IO.File.Create(tempFile))
            {
                await wordDoc.CopyToAsync(stream);
            }

            // Create the output folder.
            string moduleFolder = Path.GetFileNameWithoutExtension(filename);
            string outputPath = Path.Combine(baseFolder, moduleFolder);
            Directory.CreateDirectory(outputPath);

            try
            {
                _logger.LogDebug($"DocxToLearn(inputFile:{tempFile}, outputPath:{outputPath})");
                await DocxToLearn.ConvertAsync(tempFile, outputPath, false, useLegacyConverter
                    ? DocxConverterFactory.WithPandoc
                    : DocxConverterFactory.WithDxPlus);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.ToString());
                ModelState.AddModelError(string.Empty, ex.Message);
                return View(nameof(Index));
            }

            string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
            if (System.IO.File.Exists(zipFile))
            {
                _logger.LogDebug($"DELETE {zipFile}");
                System.IO.File.Delete(zipFile);
            }
            _logger.LogDebug($"ZIP {outputPath} => {zipFile}");
            System.IO.Compression.ZipFile.CreateFromDirectory(outputPath, zipFile);

            // Delete the temp stuff.
            _logger.LogDebug($"RMDIR {outputPath}");
            Directory.Delete(outputPath, true);
            _logger.LogDebug($"DELETE {tempFile}");
            System.IO.File.Delete(tempFile);

            if (!System.IO.File.Exists(zipFile)) return View(nameof(Index));

            // Send back the zip file.
            var fs = new FileStream(zipFile, FileMode.Open, FileAccess.Read);
            Response.RegisterForDispose(new TempFileRemover(_logger, fs, zipFile));
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
