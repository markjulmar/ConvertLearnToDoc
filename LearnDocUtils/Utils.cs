using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class Utils
    {
        private const string WinDownload = "windows-x86_64";
        private const string MacDownload = "macOS";
        private const string LatestPandocVersion = "2.16.2";

        private static async Task DownloadPandoc()
        {
            // https://github.com/jgm/pandoc/releases/download/2.16.2/pandoc-2.16.2-macOS.zip
            // https://github.com/jgm/pandoc/releases/download/2.16.2/pandoc-2.16.2-windows-x86_64.zip

            string os = RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
                ? MacDownload : WinDownload;

            string downloadUrl = $"https://github.com/jgm/pandoc/releases/download/{LatestPandocVersion}/pandoc-{LatestPandocVersion}-{os}.zip";
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

            if (os == MacDownload)
            {
                string pandocExe = PanDocExe;
                var process = Process.Start(new ProcessStartInfo
                {
                    FileName = "chmod",
                    Arguments = $"+x \"{pandocExe}\"",
                    CreateNoWindow = true
                });

                if (process != null)
                {
                    await process.WaitForExitAsync();
                }
            }
        }

        public static void CompressFolder(string folder, string zipFile) =>
            System.IO.Compression.ZipFile.CreateFromDirectory(folder, zipFile);

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
                    throw new Exception("Failed to locate runtime bin folder.");

                return Path.Combine(folder, $"pandoc-{LatestPandocVersion}",
                    RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "bin/pandoc" : "pandoc.exe");
            }
        }

        public static async Task ConvertFileAsync(Action<string> logger, string inputFile, string outputFile, string workingFolder, params string[] arguments)
        {
            const int timeout = 90; // wait up to 90s
            string executable = PanDocExe;
            if (!File.Exists(executable))
            {
                logger?.Invoke("Downloading pandoc");
                await DownloadPandoc();
            }

            if (!File.Exists(inputFile))
            {
                throw new ArgumentException($"{inputFile} does not exist.", nameof(inputFile));
            }

            if (File.Exists(outputFile))
            {
                logger?.Invoke($"DELETE {outputFile}");
                File.Delete(outputFile);
            }

            Process process;
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = executable,
                    Arguments = $"-i \"{inputFile}\" -o \"{outputFile}\" " + string.Join(' ', arguments),
                    WorkingDirectory = workingFolder,
                    CreateNoWindow = true
                };

                logger?.Invoke($"EXEC \"{processInfo.FileName} {processInfo.Arguments}\" in {processInfo.WorkingDirectory}");
                process = Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to launch {executable}", ex);
            }

            if (process == null)
                throw new Exception($"Failed to launch {executable}");

            try
            {
                await process.WaitForExitAsync(
                    new CancellationTokenSource(TimeSpan.FromSeconds(timeout)).Token);
            }
            catch (TaskCanceledException)
            {
                process.Kill(true);
                throw new Exception($"Failed to convert {inputFile} to {outputFile}, timeout after {timeout} sec.");
            }
        }
    }
}
