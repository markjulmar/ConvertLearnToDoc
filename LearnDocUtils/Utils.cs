using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class Utils
    {
        private const string LatestPandocVersion = "2.16.2";

        private static async Task DownloadPandoc()
        {
            string downloadUrl = $"https://github.com/jgm/pandoc/releases/download/{LatestPandocVersion}/pandoc-{LatestPandocVersion}-windows-x86_64.zip";
            string binFolder = BinFolder;
            string zipFile = Path.Combine(binFolder, "pandoc.zip");

            if (!File.Exists(zipFile))
            {
                using var client = new HttpClient();
                await using var downloadStream = await client.GetStreamAsync(downloadUrl);
                await using var fsOut = File.OpenWrite(zipFile);
                await downloadStream.CopyToAsync(fsOut);
            }

            System.IO.Compression.ZipFile.ExtractToDirectory(zipFile, binFolder, true);

            string pandocExe = PanDocExe;
            string folder = Path.GetDirectoryName(pandocExe);

            if (!string.IsNullOrEmpty(folder))
            {
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
            }
        }

        public static async Task<(string repo, string branch, string folder)> RetrieveLearnLocationFromUrlAsync(string moduleUrl)
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

        private static string BinFolder => Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        private static string PanDocExe
        {
            get
            {
                string folder = BinFolder;
                if (string.IsNullOrEmpty(folder))
                    throw new Exception("Failed to locate pandoc.exe");

                return Path.Combine(folder, $"pandoc-{LatestPandocVersion}", "pandoc.exe");
            }
        }

        public static async Task RunPandocAsync(string inputFile, string outputFile, string workingFolder, params string[] arguments)
        {
            string pandocExe = PanDocExe;
            if (!File.Exists(pandocExe))
            {
                await DownloadPandoc();
            }

            var process = Process.Start(new ProcessStartInfo
            {
                FileName = PanDocExe,
                Arguments = $"-i {inputFile} -o {outputFile} " + string.Join(' ', arguments),
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = workingFolder,
            });

            if (process == null)
                throw new Exception("Unable to launch pandoc.exe");

            await process.WaitForExitAsync();
        }

    }
}
