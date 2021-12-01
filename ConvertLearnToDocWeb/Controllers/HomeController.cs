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
        private readonly IConfiguration configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            this.configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ConvertLearnToDoc(LearnToDocViewModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(nameof(Index));
            }

            string repo, branch, folder;
            string outputFile = null;

            if (!string.IsNullOrEmpty(model.ModuleUrl)
                && model.ModuleUrl.ToLower().StartsWith("https"))
            {
                (repo, branch, folder) = await Utils.RetrieveLearnLocationFromUrlAsync(model.ModuleUrl);
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

            outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(folder.Split('/').Where(s => !string.IsNullOrWhiteSpace(s)).Last(), "docx"));

            try
            {
                await LearnToDocx.ConvertAsync(repo, branch, folder, outputFile, configuration.GetValue<string>("GitHub:Token"));
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, ex.Message);
            }

            if (outputFile != null && System.IO.File.Exists(outputFile))
            {
                var fs = new FileStream(outputFile, FileMode.Open, FileAccess.Read);
                Response.RegisterForDispose(new TempFileRemover(fs, outputFile));

                return File(fs, WordMimeType, Path.GetFileName(outputFile));
            }

            return View(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocToLearn(IFormFile wordDoc)
        {
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
            using (var stream = System.IO.File.Create(tempFile))
            {
                await wordDoc.CopyToAsync(stream);
            }

            // Create the output folder.
            string moduleFolder = Path.GetFileNameWithoutExtension(filename);
            string outputPath = Path.Combine(baseFolder, moduleFolder);
            Directory.CreateDirectory(outputPath);

            try
            {
                await DocxToLearn.ConvertAsync(tempFile, outputPath);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, ex.Message);
                return View(nameof(Index));
            }

            string zipFile = Path.Combine(baseFolder, Path.ChangeExtension(moduleFolder, "zip"));
            if (System.IO.File.Exists(zipFile))
            {
                System.IO.File.Delete(zipFile);
            }
            Utils.CompressFolder(outputPath, zipFile);

            // Delete the temp stuff.
            Directory.Delete(outputPath, true);
            System.IO.File.Delete(tempFile);

            // Send back the zip file.
            if (System.IO.File.Exists(zipFile))
            {
                var fs = new FileStream(zipFile, FileMode.Open, FileAccess.Read);
                Response.RegisterForDispose(new TempFileRemover(fs, zipFile));
                return File(fs, "application/zip", Path.GetFileName(zipFile));
            }

            return View(nameof(Index));
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

    internal class TempFileRemover : IDisposable
    {
        private readonly IDisposable innerDispoable;
        private readonly string path;

        public TempFileRemover(FileStream fs, string path)
        {
            this.innerDispoable = fs;
            this.path = path;
        }

        public void Dispose()
        {
            innerDispoable?.Dispose();
            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }
        }
    }
}
