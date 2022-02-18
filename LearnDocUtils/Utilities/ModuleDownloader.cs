using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
    public sealed class ModuleDownloader
    {
        public const string AbsolutePathMarker = "_fqurl_";

        public async Task<(TripleCrownModule module, string markdownFile)> DownloadModuleAsync(
            ITripleCrownGitHubService tcService, string learnFolder, string outputFolder, bool embedNotebooks)
        {
            if (string.IsNullOrEmpty(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or empty.", nameof(outputFolder));

            var module = await tcService.GetModuleAsync(learnFolder);
            if (module == null)
                throw new ArgumentException($"Failed to parse Learn module from {learnFolder}", nameof(learnFolder));

            await tcService.LoadUnitsAsync(module);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Write out index.yml
            await File.WriteAllTextAsync(Path.Combine(outputFolder, "index.yml"), module.Document);

            var markdownFile = Path.Combine(outputFolder, Path.ChangeExtension(Path.GetFileNameWithoutExtension(learnFolder)??"units.g.",".md"));
            await using var tempFile = new StreamWriter(markdownFile);

            foreach (var unit in module.Units)
            {
                // Get the unit text
                var mdText = await tcService.ReadContentForUnitAsync(unit);

                // Write out the unit file
                await WriteUnitFileAsync(unit, mdText, outputFolder);

                // Now write to our combined file
                await tempFile.WriteLineAsync($"# {unit.Title}");
                if (!string.IsNullOrEmpty(mdText))
                {
                    mdText = await DownloadAllImagesForUnit(unit.GetContentFilename(), mdText, tcService, learnFolder, outputFolder);
                    await tempFile.WriteLineAsync(mdText);
                    await tempFile.WriteLineAsync();
                }

                if (!string.IsNullOrEmpty(unit.Notebook) && embedNotebooks)
                {
                    string url = DetermineNotebookUrl(tcService, module.Url, unit.Notebook);
                    var nbText = await NotebookConverter.Convert(url);
                    if (nbText != null)
                    {
                        nbText = await DownloadAllImagesForUnit("", nbText, tcService, learnFolder, outputFolder);
                        await tempFile.WriteLineAsync(nbText);
                        await tempFile.WriteLineAsync();
                    }
                    else 
                        throw new Exception($"Failed to locate and download notebook {url}");
                }

                if (unit.Quiz != null)
                {
                    if (!string.IsNullOrEmpty(unit.Quiz.Title))
                        await tempFile.WriteLineAsync($"## {unit.Quiz.Title}{Environment.NewLine}");

                    foreach (var question in unit.Quiz.Questions)
                    {
                        await tempFile.WriteLineAsync($"### {question.Content}");
                        foreach (var choice in question.Choices)
                        {
                            await tempFile.WriteAsync(choice.IsCorrect ? "- [X]" : "- [ ]");
                            await tempFile.WriteLineAsync(EscapeContent(choice.Content));
                            if (!string.IsNullOrEmpty(choice.Explanation))
                            {
                                await tempFile.WriteLineAsync($"    - {EscapeContent(choice.Explanation)}");
                            }
                        }
                        await tempFile.WriteLineAsync();
                    }
                }
            }

            return (module, markdownFile);
        }

        private static string EscapeContent(string text) => text.Replace("{", @"\{");

        private static async Task WriteUnitFileAsync(TripleCrownUnit unit, string markdownText, string outputPath)
        {
            // Write the YAML
            string fn = Path.GetFileName(unit.Path);
            if (!string.IsNullOrEmpty(fn))
            {
                await File.WriteAllTextAsync(Path.Combine(outputPath, fn), unit.Document);
            }

            // Write the markdown
            fn = unit.GetContentFilename();
            if (!string.IsNullOrEmpty(fn)) // notebook or quiz only?
            {
                string folder = Path.GetDirectoryName(fn);
                folder = string.IsNullOrEmpty(folder) ? outputPath : Path.Combine(outputPath, folder);

                if (!string.IsNullOrEmpty(folder))
                {
                    if (!Directory.Exists(folder))
                        Directory.CreateDirectory(folder);

                    fn = Path.Combine(folder, Path.GetFileName(fn));
                }

                await File.WriteAllTextAsync(fn, markdownText);
            }
        }

        private static string DetermineNotebookUrl(ITripleCrownGitHubService service, string moduleUrl, string unitNotebook)
        {
            if (unitNotebook.ToLower().StartsWith("http"))
                return unitNotebook;

            // Look for an absolute link.
            if (unitNotebook.StartsWith('/'))
            {
                string path = "/learn/modules/";

                string[] moduleParts = moduleUrl.Split('/', '\\'); // support local and remote paths.
                path += moduleParts.Last(s => !string.IsNullOrEmpty(s));

                Debug.Assert(unitNotebook.StartsWith(path));
                unitNotebook = unitNotebook[(path.Length + 1)..];
            }

            if (!service.RootPath.ToLower().StartsWith("http")) 
                return Path.Combine(service.RootPath, unitNotebook);
            
            string url = moduleUrl;
            if (!moduleUrl.EndsWith('/'))
                url += '/';
            url += unitNotebook;
            return url;
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

        private async Task<string> DownloadAllImagesForUnit(string markdownFilename, string markdownText, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            Debug.Assert(!string.IsNullOrEmpty(markdownText));

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
                    case TripleColonBlock tcb when tcb.Extension.Name == "image":
                        tcb.Attributes.TryGetValue("source", out var bsource);
                        imageUrl = bsource;
                        break;
                }

                // Download the image.
                if (!string.IsNullOrEmpty(imageUrl) && !imageReplacements.ContainsKey(imageUrl))
                {
                    string newUrl = await DownloadImageAsync(markdownFilename, imageUrl, gitHub, moduleFolder, tempFolder);
                    if (newUrl != null)
                        imageReplacements.Add(imageUrl, newUrl);
                }
            }

            // Replace any URLs.
            return imageReplacements.Aggregate(markdownText, (current, kvp) 
                => current.Replace(kvp.Key, kvp.Value));
            
        }

        public static async Task<TripleCrownModule> GetModuleFromUidAsync(string uid)
        {
            try
            {
                var catalog = await MSLearnCatalogAPI.CatalogApi.GetCatalog();
                var url = catalog.Modules.SingleOrDefault(m => m.Uid == uid)?.BaseUrl();
                if (string.IsNullOrEmpty(url)) return null;

                var (repository, branch, folder) = await LearnResolver.LocationFromUrlAsync(url);
                var service = TripleCrownGitHubService.CreateFromToken(repository, branch);

                var module = await service.GetModuleAsync(folder);
                if (module != null)
                {
                    await service.LoadUnitsAsync(module);
                    return module;
                }
            }
            catch
            {
            }

            return null;
        }

        private static async Task<string> DownloadImageAsync(string markdownFilename, string imagePath, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            // Ignore urls.
            if (imagePath.StartsWith("http"))
                return null;

            if (Path.GetInvalidPathChars().Any(c => imagePath.Contains(c)))
                return null;

            // Remove any relative path info
            if (imagePath.StartsWith(@"../") || imagePath.StartsWith(@"..\"))
                imagePath = imagePath[3..];
            else
            {
                string includeFolder = Path.GetDirectoryName(markdownFilename)??"";
                imagePath = Path.Combine(includeFolder, imagePath);
            }

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
                    remotePath = Path.Combine(remotePath, properImagePath.TrimStart('/','\\'));
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
    }
}
