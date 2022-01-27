using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
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
            var files = new Dictionary<string, UnitMetadata>();
            using (var reader = new StreamReader(markdownFile))
            {
                UnitMetadata currentUnit = null;
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (line == null) break;

                    // Skip starting blank lines.
                    if (string.IsNullOrWhiteSpace(line) && (currentUnit == null || currentUnit.Lines.Count==0))
                        continue;

                    if (line.StartsWith("# "))
                    {
                        if (currentUnit?.Lines.Count>0)
                        {
                            // Remove any trailing empty lines.
                            while (string.IsNullOrWhiteSpace(currentUnit.Lines[^1]))
                                currentUnit.Lines.RemoveAt(currentUnit.Lines.Count-1);
                        }

                        var title = line[2..];
                        currentUnit = new UnitMetadata(title);

                        // Get a unique filename based on the title.
                        var baseFn = GenerateFilenameFromTitle(title);
                        var fn = baseFn;
                        int n = 1;
                        while (files.ContainsKey(fn))
                            fn = $"{baseFn}{n++}";

                        files.Add(fn, currentUnit);
                    }
                    else
                    {
                        currentUnit?.Lines.Add(line);
                    }
                }
            }

            var moduleUid = !string.IsNullOrEmpty(metadata.ModuleUid)
                ? metadata.ModuleUid
                : "learn." + GenerateFilenameFromTitle(metadata.Title ?? "");

            int index = 1;
            List<string> unitIds = new();

            // Write all the units.
            foreach (var (baseFn, unitMetadata) in files)
            {
                string unitFileName = $"{index}-{baseFn}";

                if (!LoadDocumentUnitMetadata(docxFile, unitMetadata))
                {
                    throw new Exception($"Failed to identify unit section in document for: \"{unitMetadata.Title}\"");
                }

                var quizText = ExtractQuiz(unitMetadata.Title, unitMetadata.Lines);
                var values = new Dictionary<string, string>
                {
                    { "module-uid", moduleUid },
                    { "unit-uid", baseFn },
                    { "title", unitMetadata.Title },
                    { "duration", EstimateDuration(unitMetadata.Lines, quizText).ToString() },
                    { "mstopic", metadata.MsTopic ?? "interactive-tutorial" },
                    { "msproduct", metadata.MsProduct ?? "learning-azure" },
                    { "msauthor", metadata.MsAuthor ?? "TBD" },
                    { "author", metadata.GitHubAlias ?? "TBD" },
                    { "interactivity", unitMetadata.BuildInteractivityOptions() },
                    { "unit-content", unitMetadata.HasContent ? $"content: |\r\n  [!include[](includes/{unitFileName}.md)]" : "" },
                    { "quizText", quizText }
                };

                string unitYaml = PopulateTemplate("unit.yml", values);
                await File.WriteAllTextAsync(Path.Combine(outputFolder, Path.ChangeExtension(unitFileName, "yml")),
                    unitYaml);

                if (unitMetadata.HasContent)
                {
                    await File.WriteAllTextAsync(Path.Combine(includeFolder, Path.ChangeExtension(unitFileName, "md")),
                        PostProcessMarkdown(string.Join("\r\n", unitMetadata.Lines)));
                }

                unitIds.Add($"- {moduleUid}.{baseFn}");
                index++;
            }

            // Write the index.yml file.
            var moduleValues = new Dictionary<string, string>
            {
                { "module-uid", moduleUid },
                { "title", metadata.Title ?? "TBD" },
                { "summary", metadata.Summary ?? "TBD" },
                { "seotitle", metadata.SEOTitle ?? "TBD" },
                { "seodescription", metadata.SEODescription ?? "TBD" },
                { "abstract", metadata.Abstract ?? "TBD" },
                { "prerequisites", metadata.Prerequisites ?? "TBD" },
                { "iconUrl", metadata.IconUrl ?? "/learn/achievements/generic-badge.svg" },
                { "saveDate", metadata.LastModified.ToString("MM/dd/yyyy") }, // 09/24/2018
                { "mstopic", metadata.MsTopic ?? "interactive-tutorial" },
                { "msproduct", metadata.MsProduct ?? "learning-azure" },
                { "msauthor", metadata.MsAuthor ?? "TBD" },
                { "author", metadata.GitHubAlias ?? "TBD" },
                { "badge-uid", metadata.BadgeUid ?? $"{moduleUid}-badge" },
                { "levels-list", metadata.GetList(metadata.Levels, "- beginner") },
                { "roles-list", metadata.GetList(metadata.Roles, "- developer") },
                { "products-list", metadata.GetList(metadata.Products, "- azure") },
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
            string result = template;
            foreach (var value in values)
            {
                string lookFor = "{" + value.Key + "}";
                if (value.Value != "")
                {
                    result = result.Replace(lookFor, value.Value.TrimEnd('\r','\n'));
                }
                else
                {
                    while (result.Contains(lookFor))
                    {
                        int start = result.IndexOf(lookFor, StringComparison.Ordinal);
                        int end = start + lookFor.Length-1;

                        while (result.Length > end+1)
                        {
                            if (result[end + 1] == '\r' || result[end + 1] == '\n')
                                end++;
                            else break;
                        }

                        result = result.Remove(start, end-start+1);
                    }
                }
            }

            return result;
        }

        private static bool LoadDocumentUnitMetadata(string docxFile, UnitMetadata unitMetadata)
        {
            var doc = Document.Load(docxFile);
            foreach (var header in doc.Paragraphs.Where(p => p.Properties.StyleName == "Heading1"))
            {
                string text = header.Text;
                if (unitMetadata.Title == text)
                {
                    var tags = string.Join(' ', header.Comments.SelectMany(c =>
                            c.Comment.Paragraphs.Select(p => p.Text)))
                        .Split(' ');
                    foreach (var tag in tags.Where(t => !string.IsNullOrEmpty(t)))
                    {
                        string value = tag.Trim().ToLower();
                        if (value == "sandbox")
                            unitMetadata.Sandbox = true;
                        else if (value.StartsWith("labid:"))
                            unitMetadata.LabId = int.Parse(value["labid:".Length..]);
                        else if (value.StartsWith("notebook:"))
                            unitMetadata.Notebook = value["notebook:".Length..].Trim();
                        if (value.StartsWith("interactivity:"))
                            unitMetadata.Interactivity = value["interactivity:".Length..].Trim();
                    }

                    return true;
                }
            }

            return false;
        }

        private static string GetProperty(IDocument doc, DocumentPropertyName name) =>
            doc.DocumentProperties.TryGetValue(name, out var text) && !string.IsNullOrWhiteSpace(text)
                ? text
                : null;

        private static string GetCustomProperty(IDocument doc, string name) =>
            doc.CustomProperties.TryGetValue(name, out var text) && text != null
                ? text.ToString()
                : null;

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

            metadata.Title ??= GetProperty(doc, DocumentPropertyName.Title);
            metadata.Summary ??= GetProperty(doc, DocumentPropertyName.Subject);
            metadata.MsAuthor ??= GetProperty(doc, DocumentPropertyName.Creator);

            // Use SaveDate first, then CreatedDate if unavailable.
            if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.SaveDate, out var dtText)
                && !string.IsNullOrWhiteSpace(dtText))
                metadata.LastModified = DateTime.Parse(dtText);
            else if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.CreatedDate, out dtText)
                     && !string.IsNullOrWhiteSpace(dtText))
                metadata.LastModified = DateTime.Parse(dtText);

            metadata.ModuleUid = GetCustomProperty(doc, nameof(ModuleMetadata.ModuleUid));
            metadata.MsTopic = GetCustomProperty(doc, nameof(ModuleMetadata.MsTopic));
            metadata.MsProduct = GetCustomProperty(doc, nameof(ModuleMetadata.MsProduct));
            metadata.Abstract = GetCustomProperty(doc, nameof(ModuleMetadata.Abstract));
            metadata.Prerequisites = GetCustomProperty(doc, nameof(ModuleMetadata.Prerequisites));
            metadata.GitHubAlias = GetCustomProperty(doc, nameof(ModuleMetadata.GitHubAlias));
            metadata.IconUrl = GetCustomProperty(doc, nameof(ModuleMetadata.IconUrl));

            metadata.Levels = GetProperty(doc, DocumentPropertyName.Comments);
            metadata.Roles = GetProperty(doc, DocumentPropertyName.Category);
            metadata.Products = GetProperty(doc, DocumentPropertyName.Keywords);
            metadata.BadgeUid = GetCustomProperty(doc, nameof(ModuleMetadata.BadgeUid));
            metadata.GitHubAlias = GetCustomProperty(doc, nameof(ModuleMetadata.GitHubAlias));
            metadata.SEOTitle = GetCustomProperty(doc, nameof(ModuleMetadata.SEOTitle));
            metadata.SEODescription = GetCustomProperty(doc, nameof(ModuleMetadata.SEODescription));

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

            // If we have an H2, then make that the title.
            // Find the start of the quiz. We're going to assume it's the first H3.
            int firstLine = -1;
            int titleLine = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                string check = lines[i].Trim().ToLower();

                if (check.StartsWith("## "))
                {
                    title = lines[i][2..].Trim();
                    titleLine = i;
                }

                else if (check.StartsWith("### "))
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

            // Make sure to remove the title since we're pulling it into the quiz.
            lines.RemoveRange(titleLine >= 0 ? titleLine : firstLine, lines.Count - (titleLine >= 0 ? titleLine : firstLine));

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
                    var isCorrect = line.Contains("❎");

                    var content = line.Replace("❎", "")
                        .Replace("⬜", "")
                        .TrimStart(' ', '-')
                        .TrimEnd();

                    sb.AppendLine($"    - content: \"{content}\"");
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