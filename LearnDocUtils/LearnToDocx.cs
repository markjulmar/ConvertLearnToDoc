using System;
using DXPlus;
using MSLearnRepos;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace LearnDocUtils
{
    public static class LearnToDocx
    {
        private static Action<string> _logger;
        private static string _accessToken;

        public static async Task ConvertAsync(string repo, string branch, string folder, string outputFile, string accessToken = null, Action<string> logger = null)
        {
            _logger = logger ?? Console.WriteLine;
            
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            _accessToken = string.IsNullOrEmpty(accessToken)
                ? GithubHelper.ReadDefaultSecurityToken()
                : accessToken;

            await Convert(TripleCrownGitHubService.CreateFromToken(repo, branch, _accessToken), folder, outputFile);
        }

        public static async Task ConvertAsync(string learnFolder, string outputFile, Action<string> logger = null)
        {
            _logger = logger ?? Console.WriteLine;

            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            await Convert(TripleCrownGitHubService.CreateLocal(learnFolder), learnFolder, outputFile);
        }

        private static async Task Convert(ITripleCrownGitHubService tcService, string learnFolder, string outputFile)
        {
            if (Directory.Exists(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' is a folder.", nameof(outputFile));

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' cannot be null or empty.", nameof(outputFile));

            if (!Path.HasExtension(outputFile))
                outputFile = Path.ChangeExtension(outputFile, "docx");

            var module = await tcService.GetModuleAsync(learnFolder);
            if (module == null)
            {
                _logger?.Invoke($"Unable to parse module from {learnFolder}.");
                return;
            }

            await tcService.LoadUnitsAsync(module);

            if (Path.GetDirectoryName(outputFile) == string.Empty)
            {
                outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), outputFile);
            }

            _logger?.Invoke($"Converting \"{module.Title}\" to {outputFile}");

            var rootTemp = Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            var includeFolder = Path.Combine(tempFolder, "includes");

            try
            {
                _logger?.Invoke($"MKDIR {includeFolder}");
                Directory.CreateDirectory(includeFolder);
                var localMarkdown = Path.Combine(includeFolder, "generated-temp.md");

                await using (var tempFile = new StreamWriter(localMarkdown))
                {
                    await tempFile.WriteLineAsync("---");
                    await tempFile.WriteLineAsync($"title: {module.Title}");
                    await tempFile.WriteLineAsync("---");
                    await tempFile.WriteLineAsync();

                    foreach (var unit in module.Units)
                    {
                        await tempFile.WriteLineAsync($"# {unit.Title}");
                        var text = await tcService.ReadContentForUnitAsync(unit);
                        if (text != null)
                        {
                            await tempFile.WriteLineAsync(ProcessMarkdownText(text));
                            await tempFile.WriteLineAsync();
                            await DownloadAllImagesForUnit(text, tcService, learnFolder, tempFolder);
                        }

                        if (unit.Quiz != null)
                        {
                            if (text == null)
                                await tempFile.WriteLineAsync($"## {unit.Quiz.Title}\r\n");

                            foreach (var question in unit.Quiz.Questions)
                            {
                                await tempFile.WriteLineAsync($"### {question.Content}");
                                foreach (var choice in question.Choices)
                                {
                                    await tempFile.WriteAsync(choice.IsCorrect ? "- [X] " : "- [ ]");
                                    await tempFile.WriteLineAsync(choice.Content);
                                    await tempFile.WriteLineAsync();
                                }
                            }

                            await tempFile.WriteLineAsync();
                        }
                    }
                }

                // Convert the file.
                await Utils.ConvertFileAsync(_logger, localMarkdown, outputFile, Path.GetDirectoryName(localMarkdown),
                    "-f markdown-fenced_divs", "-t docx");

                // Do some post processing
                _logger?.Invoke($"Processing {outputFile}");
                using var doc = Document.Load(outputFile);

                // Add the metadata
                doc.SetPropertyValue(DocumentPropertyName.Creator, module.Metadata.MsAuthor);
                doc.SetPropertyValue(DocumentPropertyName.Subject, module.Summary);
                doc.AddCustomProperty("ModuleUid", module.Uid);
                /*
                TODO: fix bug
                doc.AddCustomProperty("MsTopic", module.Metadata.MsTopic);
                doc.AddCustomProperty("MsProduct", module.Metadata.MsProduct);
                doc.AddCustomProperty("Abstract", module.Abstract);
                */

                List<Paragraph> captions = new();
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                    var paragraph = doc.Paragraphs[i];

                    // Go through and add highlights to all custom Markdown extensions.
                    foreach (var (_, _) in paragraph.FindPattern(new Regex(":::(.*?):::")))
                    {
                        paragraph.Runs.ToList()
                            .ForEach(run => run.AddFormatting(new Formatting {Highlight = Highlight.Yellow}));
                    }

                    captions.AddRange(from _ in paragraph.Pictures
                        where paragraph.Runs.Count() == 1
                        select doc.Paragraphs[i + 1]);
                }

                captions.ForEach(p => p.SetText(string.Empty));

                doc.Save();
                doc.Close();
            }
            finally
            {
                _logger?.Invoke($"RMDIR {tempFolder}");
                Directory.Delete(tempFolder, true);
            }
        }

        private static async Task DownloadAllImagesForUnit(string markdownText, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            foreach (Match match in Regex.Matches(markdownText, @"!\[(.*?)\]\((.*?)\)"))
            {
                string imagePath = match.Groups[2].Value;

                if (imagePath.StartsWith(@"../") || imagePath.StartsWith(@"..\"))
                    imagePath = imagePath[3..];

                string remotePath = moduleFolder + "/" + imagePath;
                string localPath = Path.Combine(tempFolder, imagePath);

                string localFolder = Path.GetDirectoryName(localPath);
                if (!string.IsNullOrEmpty(localFolder))
                {
                    if (!Directory.Exists(localFolder))
                        Directory.CreateDirectory(localFolder);
                }

                try
                {
                    _logger?.Invoke($"Download {remotePath} => {localPath}");
                    var result = await gitHub.ReadFileForPathAsync(remotePath);
                    if (result.binary != null)
                    {
                        await File.WriteAllBytesAsync(localPath, result.binary);
                    }
                }
                catch (Octokit.ForbiddenException)
                {
                    // Image > 1Mb in size, switch to the Git Data API and download based on the sha.
                    var remote = (IRemoteTripleCrownGitHubService) gitHub;
                    await GitHelper.GetAndWriteBlobAsync(Constants.Organization,
                        gitHub.Repository, remotePath, localPath, _accessToken,remote.Branch);
                }
            }
        }

        private static string ProcessMarkdownText(string text)
        {
            text = Regex.Replace(text, @"\[!include\[(.*?)\]\((.*?)\)]", m => $"#include \"{m.Groups[2].Value}\"");
            text = Regex.Replace(text, @"<rgn>(.*?)</rgn>", _ => "@@rgn@@");

            return text;
        }
    }
}


