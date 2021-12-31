using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class PandocUtils
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

            if (File.Exists(zipFile))
                File.Delete(zipFile);

            using (var client = new HttpClient())
            {
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
                logger?.Invoke($"Downloading pandoc {LatestPandocVersion}");
                await DownloadPandoc();
            }

            if (!File.Exists(inputFile))
                throw new ArgumentException($"{inputFile} does not exist.", nameof(inputFile));

            if (File.Exists(outputFile))
                File.Delete(outputFile);

            var processInfo = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = $"-i \"{inputFile}\" -o \"{outputFile}\" " + string.Join(' ', arguments),
                WorkingDirectory = workingFolder,
                CreateNoWindow = true
            };

            Process process;
            try
            {
                process = Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error - {ex.Message}: \"{processInfo.FileName} {processInfo.Arguments}\" in {processInfo.WorkingDirectory}\"", ex);
            }

            if (process == null)
                throw new Exception($"Failed to launch \"{processInfo.FileName} {processInfo.Arguments}\" in {processInfo.WorkingDirectory}\"");

            try
            {
                await process.WaitForExitAsync(
                    new CancellationTokenSource(TimeSpan.FromSeconds(timeout)).Token);
            }
            catch (TaskCanceledException)
            {
                process.Kill(true);
                throw new Exception($"Failed to convert {inputFile} to {outputFile}, timeout after {timeout} sec. Command was \"{processInfo.FileName} {processInfo.Arguments}\" in {processInfo.WorkingDirectory}\".");
            }
        }
    }
}
