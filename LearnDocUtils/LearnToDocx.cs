using System;
using DXPlus;
using MSLearnRepos;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LearnDocUtils
{
    public sealed class LearnToDocx
    {
        private Action<string> _logger;
        private bool _keepTempFiles;
        private string _accessToken;

        public async Task ConvertAsync(string repo, string branch, string folder,
            string outputFile, string accessToken = null, Action<string> logger = null, bool debug = false)
        {
            _logger = logger ?? Console.WriteLine;
            _keepTempFiles = debug;
            
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

        public async Task ConvertAsync(string learnFolder, string outputFile, Action<string> logger = null, bool debug = false)
        {
            _logger = logger ?? Console.WriteLine;
            _keepTempFiles = debug;

            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            await Convert(TripleCrownGitHubService.CreateLocal(learnFolder), learnFolder, outputFile);
        }

        private async Task Convert(ITripleCrownGitHubService tcService, string learnFolder, string outputFile)
        {
            if (Directory.Exists(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' is a folder.", nameof(outputFile));

            if (string.IsNullOrEmpty(outputFile))
                throw new ArgumentException($"'{nameof(outputFile)}' cannot be null or empty.", nameof(outputFile));

            if (!Path.HasExtension(outputFile))
                outputFile = Path.ChangeExtension(outputFile, "docx");

            var module = await tcService.GetModuleAsync(learnFolder);
            if (module == null)
                throw new ArgumentException($"Failed to parse Learn module from {learnFolder}", nameof(learnFolder));

            await tcService.LoadUnitsAsync(module);

            if (Path.GetDirectoryName(outputFile) == string.Empty)
            {
                outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), outputFile);
            }

            _logger?.Invoke($"Converting \"{module.Title}\" to {outputFile}");

            var rootTemp = _keepTempFiles ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            var includeFolder = Path.Combine(tempFolder, "includes");

            try
            {
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
                            text = PreprocessMarkdownText(text);

                            await tempFile.WriteLineAsync(text);
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
                PostProcessDocument(module, outputFile);
            }
            finally
            {
                if (!_keepTempFiles)
                {
                    Directory.Delete(tempFolder, true);
                }
                else
                {
                    _logger?.Invoke($"Downloaded module folder: {tempFolder}");
                }
            }
        }

        private static string PreprocessMarkdownText(string text)
        {
            text = Regex.Replace(text, @"\[!include\[(.*?)\]\((.*?)\)]", m => $"#include \"{m.Groups[2].Value}\"", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<rgn>(.*?)</rgn>", _ => "@@rgn@@", RegexOptions.IgnoreCase);
            text = ConvertTripleColonImagesToTags(text);
            //text = ConvertVideoTags(text);
            //text = ConvertTripleColonTables(text);

            return text;
        }

        private static void PostProcessDocument(TripleCrownModule module, string outputFile)
        {
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
                        .ForEach(run => run.AddFormatting(new Formatting { Highlight = Highlight.Yellow }));
                }

                captions.AddRange(from _ in paragraph.Pictures
                                  where paragraph.Runs.Count() == 1
                                  select doc.Paragraphs[i + 1]);
            }

            captions.ForEach(p => p.SetText(string.Empty));

            doc.Save();
            doc.Close();
        }

        private async Task DownloadAllImagesForUnit(string markdownText, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            foreach (Match match in Regex.Matches(markdownText, @"!\[(.*?)\]\((.*?)\)"))
            {
                string imagePath = match.Groups[2].Value;
                await DownloadImageAsync(imagePath, gitHub, moduleFolder, tempFolder);
            }

            foreach (Match match in Regex.Matches(markdownText, @"<img.+src=(?:\""|\')(.+?)(?:\""|\')(?:.+?)\>"))
            {
                string imagePath = match.Groups[1].Value;
                await DownloadImageAsync(imagePath, gitHub, moduleFolder, tempFolder);
            }
        }

        private async Task DownloadImageAsync(string imagePath, ITripleCrownGitHubService gitHub, string moduleFolder, string tempFolder)
        {
            // Ignore absolute urls.
            if (imagePath.StartsWith("http"))
                return;

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
                var (binary, _) = await gitHub.ReadFileForPathAsync(remotePath);
                if (binary != null)
                {
                    await File.WriteAllBytesAsync(localPath, binary);
                }
                else
                {
                    throw new Exception($"{remotePath} did not return an image as expected.");
                }
            }
            catch (Octokit.ForbiddenException)
            {
                // Image > 1Mb in size, switch to the Git Data API and download based on the sha.
                var remote = (IRemoteTripleCrownGitHubService)gitHub;
                await GitHelper.GetAndWriteBlobAsync(Constants.Organization,
                    gitHub.Repository, remotePath, localPath, _accessToken, remote.Branch);
            }
        }

        private static string ConvertTripleColonTables(string text)
        {
            int rows = 0;
            foreach (var (rowStart, rowEnd, rowBlock, columns) in EnumerateBoundedBlock(text, ":::row:::", ":::row-end:::"))
            {
                rows++;
                var sb = new StringBuilder("|");
                int count = 0;
                foreach (var (colStart, colEnd, colBlock, content) in EnumerateBoundedBlock(columns, ":::column:::", ":::column-end:::"))
                {
                    count++;
                    sb.Append(content.TrimEnd(' ').TrimEnd('\r').TrimEnd('\n'))
                      .Append(" |");
                }

                if (count == 0) sb.Append('|');

                if (rows == 1)
                {
                    sb.AppendLine();
                    sb.Append("|-|");
                    for (int i = 1; i < count; i++)
                    {
                        sb.Append("-|");
                    }
                }

                text = text.Replace(rowBlock, sb.ToString());
            }

            return text;
        }

        private static IEnumerable<(int start, int end, string block, string innerBlock)> EnumerateBoundedBlock(string text, string startText, string endText)
        {
            int index = text.IndexOf(startText, StringComparison.InvariantCultureIgnoreCase);
            while (index >= 0)
            {
                int end = text.IndexOf(endText, index + startText.Length);
                if (end > index)
                {
                    int innerStart = index + startText.Length;
                    int innerEnd = end;
                    end += endText.Length;

                    yield return (index, end, text[index..end], text[innerStart..innerEnd].TrimStart('\r').TrimStart('\n'));
                    index = text.IndexOf(startText, end, StringComparison.InvariantCultureIgnoreCase);
                }
            }
        }

        private static string ConvertVideoTags(string text)
        {
            /*
            var matches = Regex.Matches(text, @">[ ]*\[!VIDEO (.*?)\]", RegexOptions.IgnoreCase);
            foreach (Match m in matches)
            {
                string url = m.Groups[1].Value;
                text = text.Replace(m.Value,
                    $"<video width=\"640\" height=\"480\" controls>\r\n\t<source src=\"{url}\" type=\"video/mp4\">\r\n</video>\r\n");
            }
            */

            return text;
        }

        private static string ConvertTripleColonImagesToTags(string text)
        {
            var matches = Regex.Matches(text, @":::image (.*)[^:::]");

            foreach (Match m in matches)
            {
                var match = new StringBuilder(m.Value);
                match = match.Replace(":::image ", "<img ")
                             .Replace("source=", "src=")
                             .Replace("alt-text=", "alt=")
                             .Replace(":::", ">");
                text = text.Replace(m.Value, match.ToString());
            }

            text = text.Replace(":::image-end:::", string.Empty);

            // Replace all raw image tags.
            matches = Regex.Matches(text, @"<img([\w\W]+?)[\/]?>");
            foreach (Match m in matches)
            {
                string values = m.Value;
                string src = GetQuotedText(values, "src");
                string alt = GetQuotedText(values, "alt");
                string modifiers = "";

                /*
                string width = GetQuotedText(values, "width");
                string height = GetQuotedText(values, "height");
                string id = GetQuotedText(values, "id");

                if (width != null || height != null || id != null)
                {
                    modifiers = " { ";
                    if (id != null)
                        modifiers += $"#{id.Trim()} ";
                    if (width != null)
                        modifiers += $"width={width.Trim()} ";
                    if (height != null)
                        modifiers += $"height={height.Trim()} ";
                    modifiers += "}";
                }
                */

                text = text.Replace(m.Value, $"![{alt}]({src}){modifiers}");
            }

            return text;
        }

        private static string GetQuotedText(string text, string value)
        {
            Match match = Regex.Match(text, @$"{value}=(?:\""|\')(.+?)(?:\""|\')", RegexOptions.IgnoreCase);
            string result = match.Groups[1]?.Value;
            return string.IsNullOrEmpty(result) ? null : result;
        }
    }
}