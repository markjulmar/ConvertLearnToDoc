using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Markdig;
using Markdig.Renderer.Docx;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;
using MSLearnRepos;

namespace LearnDocUtils
{
    public sealed class LearnUtilities
    {
        public const string AbsolutePathMarker = "_fqurl_";

        public async Task<(TripleCrownModule module, string markdownFile)> DownloadModuleAsync(
            ITripleCrownGitHubService tcService, string token,
            string learnFolder, string outputFolder)
        {
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or empty.", nameof(outputFolder));

            var module = await tcService.GetModuleAsync(learnFolder);
            if (module == null)
                throw new ArgumentException($"Failed to parse Learn module from {learnFolder}", nameof(learnFolder));

            await tcService.LoadUnitsAsync(module);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var markdownFile = Path.Combine(outputFolder, Path.ChangeExtension(Path.GetFileNameWithoutExtension(learnFolder)??"temp-markdown",".md"));

            await using var tempFile = new StreamWriter(markdownFile);

            foreach (var unit in module.Units)
            {
                await tempFile.WriteLineAsync($"# {unit.Title}");
                var text = await tcService.ReadContentForUnitAsync(unit);
                if (text != null)
                {
                    text = await DownloadAllImagesForUnit(text, tcService, learnFolder, outputFolder);
                    await tempFile.WriteLineAsync(text);
                    await tempFile.WriteLineAsync();
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

            return (module, markdownFile);
        }

        private readonly Lazy<MarkdownPipeline> markdownPipeline = new(CreatePipeline);
        private static MarkdownPipeline CreatePipeline()
        {
            var context = new MarkdownContext();
            var pipelineBuilder = new MarkdownPipelineBuilder();
            return pipelineBuilder
                .UsePipeTables()
                .UseRow(context)
                .UseNestedColumn(context)
                .UseTripleColon(context)
                .UseGenericAttributes() // Must be last as it is one parser that is modifying other parsers
                .Build();
        }

        private async Task<string> DownloadAllImagesForUnit(string markdownText, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            var markdownDocument = Markdown.Parse(markdownText, markdownPipeline.Value);
            Dictionary<string, string> imageReplacements = new();

            foreach (var item in markdownDocument.EnumerateBlocks())
            {
                string imageUrl = null;

                switch (item)
                {
                    case LinkInline { IsImage: true } li:
                        imageUrl = li.Url;
                        break;
                    case HtmlBlock html:
                    {
                        var tag = html.Lines.ToString();
                        var result = Regex.Match(tag, @"<img.*\s+src=(?:\""|\')(.+?)(?:\""|\')(?:.+?)\>");
                        if (result.Success)
                        {
                            imageUrl = result.Groups[1].Value;
                        }
                        break;
                    }
                    case TripleColonInline tci when tci.Extension.Name == "image":
                        tci.Attributes.TryGetValue("source", out var source);
                        imageUrl = source;
                        break;
                }

                // Download the image.
                if (!string.IsNullOrEmpty(imageUrl) && !imageReplacements.ContainsKey(imageUrl))
                {
                    string newUrl = await DownloadImageAsync(imageUrl, gitHub, moduleFolder, tempFolder);
                    if (newUrl != null)
                        imageReplacements.Add(imageUrl, newUrl);
                }
            }

            // Replace any URLs.
            return imageReplacements.Aggregate(markdownText, (current, kvp) 
                => current.Replace(kvp.Key, kvp.Value));
            
        }

        private static async Task<string> DownloadImageAsync(string imagePath, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            // Ignore urls.
            if (imagePath.StartsWith("http"))
                return null;

            if (Path.GetInvalidPathChars().Any(c => imagePath.Contains(c)))
                return null;

            // Remove any relative path info
            if (imagePath.StartsWith(@"../") || imagePath.StartsWith(@"..\"))
                imagePath = imagePath[3..];

            string properImagePath = imagePath;
            int index = properImagePath.IndexOf('?');
            if (index > 0)
                properImagePath = properImagePath[..index];
            index = properImagePath.IndexOf('#');
            if (index > 0)
                properImagePath = properImagePath[..index];

            var rgh = gitHub as IRemoteTripleCrownGitHubService;

            bool absolutePath = false;
            string remotePath;

            if (properImagePath.StartsWith('/'))
            {
                absolutePath = true;

                // Pull out the composite path. Should start with "/learn"
                if (properImagePath.ToLower().StartsWith("/learn"))
                    properImagePath = properImagePath["/learn".Length..];

                if (rgh != null)
                {
                    Uri uri = new Uri(gitHub.RootPath, UriKind.Absolute);
                    remotePath =  uri.Segments.Last() + properImagePath;
                }
                else // local
                {
                    string rootFolder = gitHub.RootPath;
                    string[] parts = rootFolder.Split(new[] { '\\', '/' });

                    int lastMatch = -1;
                    for (var i = 0; i < parts.Length; i++)
                    {
                        var item = parts[i];
                        if (Regex.Match(item, "learn(.*)-pr").Success)
                        {
                            lastMatch = i;
                        }
                    }
                    remotePath = lastMatch == -1 ? rootFolder : string.Join(Path.DirectorySeparatorChar, parts.Take(lastMatch + 1));
                }
            }
            else
            {
                remotePath = moduleFolder + "/" + properImagePath;
            }

            string localPath;
            if (absolutePath)
            {
                imagePath = Path.Combine(AbsolutePathMarker, imagePath.TrimStart('/', '\\')).Replace('\\', '/');
                localPath = Path.Combine(tempFolder, imagePath);
            }
            else
            {
                localPath = Path.Combine(tempFolder, properImagePath);
            }

            // Already downloaded?
            if (File.Exists(localPath))
                return imagePath;

            string localFolder = Path.GetDirectoryName(localPath);
            if (!string.IsNullOrEmpty(localFolder))
            {
                if (!Directory.Exists(localFolder))
                    Directory.CreateDirectory(localFolder);
            }

            try
            {
                var (binary, _) = await gitHub.ReadFileForPathAsync(remotePath);
                if (binary != null)
                {
                    await File.WriteAllBytesAsync(localPath, binary);
                }
                else
                {
                    throw new Exception($"\"{remotePath}\" did not return an image as expected.");
                }
            }
            catch (Octokit.ForbiddenException)
            {
                // Image > 1Mb in size, switch to the Git Data API and download based on the sha.
                await rgh!.Helper.GetAndWriteBlobAsync( remotePath, localPath, rgh.Branch);
            }

            return imagePath;
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
    }
}
