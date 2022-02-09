using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DXPlus;
using MSLearnRepos;

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
            var metadata = await LoadDocumentMetadata(docxFile);

            // Get the UID
            var moduleUid = !string.IsNullOrEmpty(metadata.ModuleData.Uid)
                ? metadata.ModuleData.Uid
                : "learn." + GenerateFilenameFromTitle(metadata.ModuleData.Title ?? "");

            string includeFolder = Path.Combine(outputFolder, "includes");
            Directory.CreateDirectory(includeFolder);

            // Read all the lines
            int index = 1;
            var files = new Dictionary<string, UnitMetadata>();
            using (var reader = new StreamReader(markdownFile))
            {
                bool inCodeBlock = false;
                UnitMetadata currentUnit = null;
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (line == null) break;

                    // Skip starting blank lines.
                    if (string.IsNullOrWhiteSpace(line) && (currentUnit == null || currentUnit.Lines.Count==0))
                        continue;

                    if (line.Trim().StartsWith("```"))
                    {
                        inCodeBlock = !inCodeBlock;
                    }

                    if (!inCodeBlock && line.StartsWith("# "))
                    {
                        if (currentUnit?.Lines.Count>0)
                        {
                            // Remove any trailing empty lines.
                            while (string.IsNullOrWhiteSpace(currentUnit.Lines[^1]))
                                currentUnit.Lines.RemoveAt(currentUnit.Lines.Count-1);
                        }

                        var title = line[2..];

                        // See if we can identify the unit metadata.
                        TripleCrownUnit unitMetadata = null;
                        int pos = files.Count;
                        if (metadata.ModuleData.Units.Count > pos)
                        {
                            var umd = metadata.ModuleData.Units[pos];
                            if (FuzzyCompare(umd.Title, title) > .9)
                                unitMetadata = umd;
                        }

                        unitMetadata ??= new TripleCrownUnit
                        {
                            Title = title,
                            Metadata = new MSLearnRepos.UnitMetadata
                            {
                                Title = title,
                                Author = metadata.ModuleData.Metadata.Author,
                                MsAuthor = metadata.ModuleData.Metadata.MsAuthor,
                                MsDate = metadata.ModuleData.Metadata.MsDate,
                            }
                        };

                        currentUnit = new UnitMetadata(title, unitMetadata);

                        // Get the original filename if we can determine what it was.
                        string fn = unitMetadata.GetContentFilename();
                        if (string.IsNullOrEmpty(fn))
                        {
                            // Get a unique filename based on the title + unit index.
                            var baseFn = GenerateFilenameFromTitle(title);
                            if (string.IsNullOrEmpty(unitMetadata.Uid))
                            {
                                // Account for repeated titles.
                                var uid = $"{moduleUid}.{baseFn}";
                                int suffix = 1;
                                while (files.Any(f => f.Value.Metadata.Uid == baseFn))
                                {
                                    uid = $"{moduleUid}.{baseFn}{suffix}";
                                    suffix++;
                                }
                                unitMetadata.Uid = uid;
                            }
                            fn = $"{index}-{baseFn}";
                        }
                        else
                        {
                            fn = Path.ChangeExtension(Path.GetFileName(fn), ".md");
                        }

                        files.Add(fn, currentUnit);
                        index++;
                    }
                    else
                    {
                        currentUnit?.Lines.Add(line);
                    }
                }
            }

            List<string> unitIds = new();

            // Write all the units.
            foreach (var (unitFileName, unitMetadata) in files)
            {
                if (!LoadDocumentUnitMetadata(docxFile, unitMetadata))
                    continue;

                var quizText = ExtractQuiz(unitMetadata.Title, unitMetadata.Lines);

                var zonePivotGroups = !string.IsNullOrEmpty(unitMetadata.Metadata.Metadata.ZonePivotGroups)
                    ? "  zone_pivot_groups: " + unitMetadata.Metadata.Metadata.ZonePivotGroups
                    : null;

                var tasks = BuildTaskValidation(unitMetadata.Metadata.Tasks);

                var values = new Dictionary<string, string>
                {
                    { "module-uid", moduleUid },
                    { "unit-uid", unitMetadata.Metadata.Uid },
                    { "title", unitMetadata.Title },
                    { "seotitle", unitMetadata.Metadata.Metadata.Title ?? unitMetadata.Title },
                    { "seodescription", unitMetadata.Metadata.Metadata.Description ?? "TBD" },
                    { "duration", EstimateDuration(unitMetadata.Metadata.DurationInMinutes, unitMetadata.Lines, quizText).ToString() },
                    { "saveDate", metadata.ModuleData.Metadata.MsDate }, // 09/24/2018
                    { "mstopic", unitMetadata.Metadata.Metadata.MsTopic ?? "interactive-tutorial" },
                    { "msproduct", unitMetadata.Metadata.Metadata.MsProduct ?? "learning-azure" },
                    { "msauthor", unitMetadata.Metadata.Metadata.MsAuthor ?? "TBD" },
                    { "author", unitMetadata.Metadata.Metadata.Author ?? "TBD" },
                    { "interactivity", unitMetadata.BuildInteractivityOptions() },
                    { "unit-content", unitMetadata.HasContent ? $"content: |\r\n  [!include[](includes/{unitFileName})]" : "" },
                    { "zonePivots", zonePivotGroups },
                    { "quizText", quizText },
                    { "task-validation", tasks }
                };

                string unitYaml = PopulateTemplate("unit.yml", values);
                await File.WriteAllTextAsync(Path.Combine(outputFolder, Path.ChangeExtension(unitFileName, "yml")),
                    unitYaml);

                if (unitMetadata.HasContent)
                {
                    await File.WriteAllTextAsync(Path.Combine(includeFolder, Path.ChangeExtension(unitFileName, "md")),
                        PostProcessMarkdown(string.Join("\r\n", unitMetadata.Lines)));
                }

                unitIds.Add($"- {unitMetadata.Metadata.Uid}");
            }

            // Write the index.yml file.
            var moduleValues = new Dictionary<string, string>
            {
                { "module-uid", moduleUid },
                { "title", metadata.ModuleData.Title ?? "TBD" },
                { "summary", FormatMultiLineYaml(metadata.ModuleData.Summary) ?? "TBD" },
                { "seotitle", metadata.ModuleData.Metadata.Title ?? "TBD" },
                { "seodescription", metadata.ModuleData.Metadata.Description ?? "TBD" },
                { "abstract", FormatMultiLineYaml(metadata.ModuleData.Abstract) ?? "TBD" },
                { "prerequisites", FormatMultiLineYaml(metadata.ModuleData.Prerequisites) ?? "TBD" },
                { "iconUrl", metadata.ModuleData.IconUrl ?? "/learn/achievements/generic-badge.svg" },
                { "saveDate", metadata.ModuleData.Metadata.MsDate }, // 09/24/2018
                { "mstopic", metadata.ModuleData.Metadata.MsTopic ?? "interactive-tutorial" },
                { "msproduct", metadata.ModuleData.Metadata.MsProduct ?? "learning-azure" },
                { "msauthor", metadata.ModuleData.Metadata.MsAuthor ?? "TBD" },
                { "author", metadata.ModuleData.Metadata.Author ?? "TBD" },
                { "badge-uid", metadata.ModuleData.Badge?.Uid ?? $"{moduleUid}-badge" },
                { "levels-list", ModuleMetadata.GetList(metadata.ModuleData.Levels, "- beginner") },
                { "roles-list", ModuleMetadata.GetList(metadata.ModuleData.Roles, "- developer") },
                { "products-list", ModuleMetadata.GetList(metadata.ModuleData.Products, "- azure") },
                { "unit-uid-list", string.Join("\r\n", unitIds) }
            };

            await File.WriteAllTextAsync(Path.Combine(outputFolder, "index.yml"),
                PopulateTemplate("index.yml", moduleValues));
        }

        private string BuildTaskValidation(List<TripleCrownTaskValidation> tasks)
        {
            if (tasks == null || tasks.Count == 0) return null;

            var sb = new StringBuilder().AppendLine("tasks:");
            foreach (var task in tasks)
            {
                sb.AppendLine($"- action: {task.Action}")
                  .AppendLine($"  environment: {task.Environment}");
                if (!string.IsNullOrEmpty(task.Hint))
                    sb.AppendLine($"  hint: \"{task.Hint}\"");
                if (task.Azure != null)
                {
                    sb.AppendLine("  azure:");

                    if (!string.IsNullOrEmpty(task.Azure.ResourceGroup))
                        sb.AppendLine($"    resourceGroup: {task.Azure.ResourceGroup}");

                    if (task.Azure.Tags?.Count > 0)
                    {
                        sb.AppendLine("    tags:");
                        foreach (var tag in task.Azure.Tags)
                        {
                            sb.AppendLine($"    - name: {tag.Name}");
                            sb.AppendLine($"      value: {tag.Value}");
                        }
                    }

                    if (task.Azure.Resource?.Count > 0)
                    {
                        sb.AppendLine("    resource:");
                        foreach (var kvp in task.Azure.Resource)
                        {
                            sb.AppendLine($"      {kvp.Key}: \"{kvp.Value}\"");
                        }
                    }
                }
            }

            return sb.ToString();
        }

        private static double FuzzyCompare(string input1, string input2)
        {
            HashSet<string> BuildBigramSet(string input)
            {
                var bigrams = new HashSet<string>();
                for (int i = 0; i < input.Length - 1; i++) bigrams.Add(input.Substring(i, 2));
                return bigrams;
            }

            var nx = BuildBigramSet(input1.ToLower());
            var ny = BuildBigramSet(input2.ToLower());

            var intersection = new HashSet<string>(nx);
            intersection.IntersectWith(ny);

            double dbOne = intersection.Count;
            return 2 * dbOne / (nx.Count + ny.Count);
        }

        private static string FormatMultiLineYaml(string text) 
            => text==null ? null : text.Contains('\n') ? "|" + Environment.NewLine + text : text;

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
                if (!string.IsNullOrEmpty(value.Value))
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
                            unitMetadata.Metadata.UsesSandbox = true;
                        else if (value.StartsWith("labid:"))
                            unitMetadata.Metadata.LabId = int.Parse(value["labid:".Length..]);
                        else if (value.StartsWith("notebook:"))
                            unitMetadata.Metadata.Notebook = value["notebook:".Length..].Trim();
                        if (value.StartsWith("interactivity:"))
                            unitMetadata.Metadata.InteractivityType = value["interactivity:".Length..].Trim();
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

        private static async Task<ModuleMetadata> LoadDocumentMetadata(string docxFile)
        {
            var doc = Document.Load(docxFile);

            TripleCrownModule moduleData = null;
            if (doc.CustomProperties.TryGetValue(nameof(TripleCrownModule.Metadata), out var jsonText) && jsonText != null)
            {
                moduleData = JsonHelper.FromJson<TripleCrownModule>(jsonText.ToString());
            }

            if (moduleData == null && doc.DocumentProperties.TryGetValue(DocumentPropertyName.Comments, out string uid) 
                        && !string.IsNullOrEmpty(uid) && uid.StartsWith("learn."))
            {
                moduleData = await ModuleDownloader.GetModuleFromUidAsync(uid);
            }

            var metadata = new ModuleMetadata(moduleData);

            foreach (var item in doc.Paragraphs)
            {
                var styleName = item.Properties.StyleName;
                if (styleName == "Heading1") break;
                switch (styleName)
                {
                    case "Title":
                        metadata.ModuleData.Title = item.Text;
                        break;
                    case "Author":
                        metadata.ModuleData.Metadata.MsAuthor = item.Text;
                        break;
                    case "Abstract":
                        metadata.ModuleData.Summary = item.Text;
                        break;
                }
            }

            metadata.ModuleData.Title ??= GetProperty(doc, DocumentPropertyName.Title);
            metadata.ModuleData.Summary ??= GetProperty(doc, DocumentPropertyName.Subject);
            metadata.ModuleData.Metadata.MsAuthor ??= GetProperty(doc, DocumentPropertyName.Creator);

            // Use SaveDate first, then CreatedDate if unavailable.
            if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.SaveDate, out var dtText)
                && !string.IsNullOrWhiteSpace(dtText))
                metadata.ModuleData.Metadata.MsDate = DateTime.Parse(dtText).ToString("MM/dd/yyyy");
            else if (doc.DocumentProperties.TryGetValue(DocumentPropertyName.CreatedDate, out dtText)
                && !string.IsNullOrWhiteSpace(dtText))
                metadata.ModuleData.Metadata.MsDate = DateTime.Parse(dtText).ToString("MM/dd/yyyy");

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

        private static string GetTemplate(string templateKey)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream($"LearnDocUtils.templates.{templateKey}");
            if (stream == null)
                throw new ArgumentException($"Embedded resource {templateKey} missing from assembly", nameof(templateKey));

            using var sr = new StreamReader(stream);
            return sr.ReadToEnd();
        }

        private static string ExtractQuiz(string title, List<string> lines)
        {
            int start = 0;
            for (; start < lines.Count; start++)
            {
                // Must be an H2 with the words "Knowledge" and "Check". Can be "Check your Knowledge", etc.
                if (lines[start].StartsWith("## ")
                    && lines[start].Contains("knowledge", StringComparison.InvariantCultureIgnoreCase)
                    && lines[start].Contains("check", StringComparison.InvariantCultureIgnoreCase))
                {
                    break;
                }
            }

            if (start == lines.Count) return string.Empty;

            // If we have an H2, then make that the title.
            // Find the start of the quiz. We're going to assume it's the first H3.
            int firstLine = -1, titleLine = -1;
            for (int i = start; i < lines.Count; i++)
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
            // TODO: we assume here that the quiz is always the LAST thing in the unit. This is how
            // Learn works today, but that could change. We should probably look for a H2 separator.
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
                    // Get rid of all spaces and dashes in front of the text.
                    var text = line.Trim().TrimStart('-').TrimStart();
                    Debug.Assert(text.Length>0);

                    if (text[0] == '[') // choice?
                    {
                        var isCorrect = text.Replace(" ", "").ToLower().StartsWith("[x]");
                        start = text.IndexOf(']') + 1;
                        Debug.Assert(start>=3);

                        sb.AppendLine($"    - content: \"{text[start..].Trim()}\"");
                        sb.AppendLine($"      isCorrect: {isCorrect.ToString().ToLower()}");
                    }
                    else sb.AppendLine($"      explanation: \"{text}\"");
                }
            }

            return sb.ToString();
        }

        private static int EstimateDuration(int existingValue, IEnumerable<string> lines, string quizText) =>
            existingValue > 0
                ? existingValue
                : 1 + (lines.Sum(l => l.Split(' ').Length)
                       + quizText.Split(' ').Length) / 120;
    }
}