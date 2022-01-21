using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DXPlus;

namespace LearnDocUtils
{
    public class ModuleBuilder
    {
        private readonly string docxFile;
        private readonly string outputFolder;
        private readonly string markdownFile;

        public ModuleBuilder(string docxFile, string outputFolder, string markdownFile)
        {
            this.docxFile = docxFile;
            this.outputFolder = outputFolder;
            this.markdownFile = markdownFile;
        }

        public async Task CreateModuleAsync()
        {
            // Get the title metadata.
            var metadata = LoadDocumentMetadata(docxFile);

            string includeFolder = Path.Combine(outputFolder, "includes");
            Directory.CreateDirectory(includeFolder);

            // Read all the lines
            var files = new Dictionary<string, (string title, List<string> lines)>();
            using (var reader = new StreamReader(markdownFile))
            {
                List<string> currentFile = null;
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (line == null) break;

                    // Skip starting blank lines.
                    if (string.IsNullOrWhiteSpace(line) && (currentFile == null || currentFile.Count==0))
                        continue;

                    if (line.StartsWith("# "))
                    {
                        if (currentFile != null)
                        {
                            // Remove any trailing empty lines.
                            while (string.IsNullOrWhiteSpace(currentFile[^1]))
                                currentFile.RemoveAt(currentFile.Count-1);
                        }

                        currentFile = new List<string>();
                        var title = line[2..];

                        // Get a unique filename based on the title.
                        var baseFn = GenerateFilenameFromTitle(title);
                        var fn = baseFn;
                        int n = 1;
                        while (files.ContainsKey(fn))
                        {
                            fn = $"{baseFn}{n++}";
                        }

                        files.Add(fn, (title, currentFile));
                    }
                    else
                    {
                        currentFile?.Add(line);
                    }
                }
            }

            var moduleUid = !string.IsNullOrEmpty(metadata.ModuleUid)
                ? metadata.ModuleUid
                : "learn." + GenerateFilenameFromTitle(metadata.Title ?? "");

            int index = 1;
            List<string> unitIds = new();

            // Write all the units.
            foreach (var (baseFn, (title, lines)) in files)
            {
                string unitFileName = $"{index}-{baseFn}";

                var quizText = ExtractQuiz(title, lines);
                var values = new Dictionary<string, string>
                {
                    { "module-uid", moduleUid },
                    { "unit-uid", baseFn },
                    { "title", title },
                    { "duration", EstimateDuration(lines, quizText).ToString() },
                    { "mstopic", metadata.MsTopic ?? "interactive-tutorial" },
                    { "msproduct", metadata.MsProduct ?? "learning-azure" },
                    { "author", metadata.MsAuthor ?? "TBD" },
                    { "unit-content", $"  |\r\n  [!include[](includes/{unitFileName}.md)]" },
                    { "quizText", quizText }
                };

                string unitYaml = PopulateTemplate("unit.yml", values);
                await File.WriteAllTextAsync(Path.Combine(outputFolder, Path.ChangeExtension(unitFileName, "yml")),
                    unitYaml);
                await File.WriteAllTextAsync(Path.Combine(includeFolder, Path.ChangeExtension(unitFileName, "md")),
                    PostProcessMarkdown(string.Join("\r\n", lines)));
                unitIds.Add($"- {moduleUid}.{baseFn}");
                index++;
            }

            // Write the index.yml file.
            var moduleValues = new Dictionary<string, string>
            {
                { "module-uid", moduleUid },
                { "title", metadata.Title ?? "TBD" },
                { "summary", metadata.Summary ?? "TBD" },
                { "saveDate", metadata.LastModified.ToString("MM/dd/yyyy") }, // 09/24/2018
                { "mstopic", metadata.MsTopic ?? "interactive-tutorial" },
                { "msproduct", metadata.MsProduct ?? "learning-azure" },
                { "author", metadata.MsAuthor ?? "TBD" },
                { "unit-uid-list", string.Join("\r\n", unitIds) }
            };

            await File.WriteAllTextAsync(Path.Combine(outputFolder, "index.yml"),
                PopulateTemplate("index.yml", moduleValues));
        }

        private static string PostProcessMarkdown(string text)
        {
            text = text.Replace("(./media/", "(../media/");
            text = text.Replace("\"media/", "\"../media/");
            return text;
        }

        private static string PopulateTemplate(string templateKey, Dictionary<string, string> values)
        {
            string template = GetTemplate(templateKey);
            foreach (var item in values)
            {
                template = template.Replace("{" + item.Key + "}", item.Value);
            }

            return template;
        }

        private static ModuleMetadata LoadDocumentMetadata(string docxFile)
        {
            var doc = Document.Load(docxFile);
            var metadata = new ModuleMetadata();

            foreach (var item in doc.Paragraphs)
            {
                var styleName = item.Properties.StyleName;
                if (styleName == "Heading1") break;
                switch (styleName)
                {
                    case "Title":
                        metadata.Title = item.Text;
                        break;
                    case "Author":
                        metadata.MsAuthor = item.Text;
                        break;
                    case "Abstract":
                        metadata.Summary = item.Text;
                        break;
                }
            }

            if (metadata.Title == null
                && doc.DocumentProperties.TryGetValue(DocumentPropertyName.Title, out var title)
                && !string.IsNullOrWhiteSpace(title))
            {
                metadata.Title = title;
            }

            if (metadata.Summary == null
                && doc.DocumentProperties.TryGetValue(DocumentPropertyName.Subject, out var summary)
                && !string.IsNullOrWhiteSpace(summary))
            {
                metadata.Summary = summary;
            }

            if (metadata.MsAuthor == null
                && doc.DocumentProperties.TryGetValue(DocumentPropertyName.Creator, out var author)
                && !string.IsNullOrWhiteSpace(author))
            {
                metadata.MsAuthor = author;
            }

            if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.SaveDate, out var dtText)
                && !string.IsNullOrWhiteSpace(dtText))
            {
                metadata.LastModified = DateTime.Parse(dtText);
            }

            if (doc.CustomProperties.TryGetValue(nameof(ModuleMetadata.ModuleUid), out var uid)
                && uid != null)
            {
                metadata.ModuleUid = uid.ToString();
            }

            if (doc.CustomProperties.TryGetValue(nameof(ModuleMetadata.MsTopic), out var topic)
                && topic != null)
            {
                metadata.MsTopic = topic.ToString();
            }

            if (doc.CustomProperties.TryGetValue(nameof(ModuleMetadata.MsProduct), out var product)
                && product != null)
            {
                metadata.MsProduct = product.ToString();
            }

            if (doc.CustomProperties.TryGetValue(nameof(ModuleMetadata.Abstract), out var @abstract)
                && @abstract != null)
            {
                metadata.Abstract = @abstract.ToString();
            }

            return metadata;
        }

        private static string GenerateFilenameFromTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                throw new ArgumentException("Unable to determine title of document", nameof(title));

            return new string(
                        title
                            .Replace(" - ", "-")
                            .Replace(' ', '-')
                            .Where(ch => char.IsLetter(ch) || ch is '-')
                            .Select(char.ToLower)
                            .ToArray());
        }

        private static string GetTemplate(string templateKey) =>
            File.ReadAllText(Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? @".\",
                "templates", templateKey));

        private static string ExtractQuiz(string title, List<string> lines)
        {
            // Look for the KC title.
            if (title.Trim().ToLower() is not ("knowledge check" or "check your knowledge"))
                return string.Empty;

            // Find the start of the quiz. We're going to assume it's the first H3.
            int firstLine = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                string check = lines[i].Trim().ToLower();
                if (check.StartsWith("### "))
                {
                    firstLine = i;
                    break;
                }
            }

            // Hmm. never found it.
            if (firstLine == -1)
                return string.Empty;

            // Get the lines specific to the quiz and remove them from content.
            var quiz = new List<string>(lines.GetRange(firstLine, lines.Count - firstLine));
            lines.RemoveRange(firstLine, lines.Count - firstLine);

            // Now build the quiz.
            var sb = new StringBuilder();
            sb.AppendLine("quiz:");
            sb.AppendLine($"  title: {title}");
            sb.AppendLine("  questions:");

            foreach (var line in quiz.Where(line => !string.IsNullOrWhiteSpace(line)))
            {
                if (line.StartsWith("### "))
                {
                    sb.AppendLine($"  - content: \"{line[4..]}\"");
                    sb.AppendLine("    choices:");
                }
                else
                {
                    var isCorrect = line.Contains("[x]") || line.Contains("☒");
                    sb.AppendLine($"    - content: \"{line.Replace("☒", "")[(line.IndexOf(']') + 1)..].TrimStart()}\"");
                    sb.AppendLine("      explanation: \"\"");
                    sb.AppendLine($"      isCorrect: {isCorrect.ToString().ToLower()}");
                }
            }

            return sb.ToString();
        }

        private static int EstimateDuration(IEnumerable<string> lines, string quizText)
        {
            return 1 + (lines.Sum(l => l.Split(' ').Length)
                        + quizText.Split(' ').Length) / 120;
        }
    }
}