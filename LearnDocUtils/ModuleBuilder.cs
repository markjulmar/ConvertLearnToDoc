﻿using System;
using System.Collections.Generic;
using System.IO;
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
        private readonly Action<string> logger;

        public ModuleBuilder(string docxFile, string outputFolder, string markdownFile, Action<string> logger)
        {
            this.docxFile = docxFile;
            this.outputFolder = outputFolder;
            this.markdownFile = markdownFile;
            this.logger = logger;
        }

        public async Task CreateModuleAsync(Func<string,string> markdownProcessor)
        {
            if (markdownProcessor == null)
                markdownProcessor = s => s;

            // Get the title metadata.
            var metadata = LoadDocumentMetadata(docxFile);

            string includeFolder = Path.Combine(outputFolder, "includes");
            Directory.CreateDirectory(includeFolder);

            // Read all the lines
            var files = new Dictionary<string, List<string>>();
            using (var reader = new StreamReader(markdownFile))
            {
                List<string> currentFile = null;
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (line == null) break;
                    if (line.StartsWith("# "))
                    {
                        currentFile = new List<string>();
                        files.Add(line.Substring(2), currentFile);
                    }
                    else if (currentFile != null)
                    {
                        currentFile.Add(line);
                    }
                }
            }

            string moduleUid = !string.IsNullOrEmpty(metadata.ModuleUid) ? metadata.ModuleUid : "learn." + GenerateFilenameFromTitle(metadata.Title ?? "");

            int index = 1;
            List<string> unitIds = new();

            // Write all the units.
            foreach (var (title, value) in files)
            {
                string baseFn = GenerateFilenameFromTitle(title);
                string unitFileName = $"{index}-{baseFn}";

                var quizText = ExtractQuiz(value);
                var values = new Dictionary<string, string>
                    {
                        { "module-uid", moduleUid },
                        { "unit-uid", baseFn },
                        { "title", title },
                        { "mstopic", metadata.MsTopic??"interactive-tutorial" },
                        { "msproduct", metadata.MsProduct??"learning-azure" },
                        { "author", metadata.MsAuthor ?? "TBD" },
                        { "unit-content", $"  |\r\n  [!include[](includes/{unitFileName}.md)]" },
                        { "quizText", quizText }
                    };

                string unitYaml = PopulateTemplate("unit.yml", values);
                await File.WriteAllTextAsync(Path.Combine(outputFolder, Path.ChangeExtension(unitFileName, "yml")), unitYaml);
                await File.WriteAllTextAsync(Path.Combine(includeFolder, Path.ChangeExtension(unitFileName, "md")), markdownProcessor.Invoke(string.Join("\r\n", value)));
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
                    { "mstopic", metadata.MsTopic??"interactive-tutorial" },
                    { "msproduct", metadata.MsProduct??"learning-azure" },
                    { "author", metadata.MsAuthor ?? "TBD" },
                    { "unit-uid-list", string.Join("\r\n", unitIds) }
                };

            await File.WriteAllTextAsync(Path.Combine(outputFolder, "index.yml"), PopulateTemplate("index.yml", moduleValues));

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

        private ModuleMetadata LoadDocumentMetadata(string docxFile)
        {
            logger?.Invoke($"LoadDocumentMetadata({docxFile})");

            var doc = Document.Load(docxFile);
            var metadata = new ModuleMetadata();

            foreach (var item in doc.Paragraphs)
            {
                var styleName = item.Properties.StyleName;
                if (styleName == "Heading1") break;
                switch (styleName)
                {
                    case "Title": metadata.Title = item.Text; break;
                    case "Author": metadata.MsAuthor = item.Text; break;
                    case "Abstract": metadata.Summary = item.Text; break;
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

            title = title.Replace(" - ", " ");

            string fn = "";
            for (int i = 0; i < title.Length; i++)
            {
                char ch = title[i];
                if (ch == ' ') ch = '-';
                if (char.IsLetter(ch) || ch == '-')
                {
                    fn += char.ToLower(ch);
                }
            }
            return fn;
        }

        private static string GetTemplate(string templateKey) =>
            File.ReadAllText(Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? @".\",
                "templates", templateKey));

        private static string ExtractQuiz(List<string> lines)
        {
            var sb = new StringBuilder();

            for (int i = 0; i < lines.Count; i++)
            {
                string check = lines[i].Trim().ToLower();
                if (check == "## knowledge check"
                    || check == "## check your knowledge")
                {
                    int lastLine = i + 1;
                    while (lastLine < lines.Count)
                    {
                        if (lines[lastLine].StartsWith("## ")
                            || lines[lastLine].StartsWith("# "))
                            break;
                        lastLine++;
                    }

                    var quiz = new List<string>(lines.GetRange(i, lastLine - i));

                    sb.AppendLine("quiz:");
                    sb.AppendLine($"  title: {quiz[0][3..]}");
                    sb.AppendLine("  questions:");

                    for (int pos = 1; pos < quiz.Count; pos++)
                    {
                        string line = quiz[pos];
                        if (string.IsNullOrWhiteSpace(line))
                            continue;
                        if (line.StartsWith("### "))
                        {
                            sb.AppendLine($"  - content: \"{line.Substring(4)}\"");
                            sb.AppendLine("    choices:");
                        }
                        else
                        {
                            sb.AppendLine($"    - content: \"{line[(line.IndexOf(']') + 1)..].TrimStart()}\"");
                            sb.AppendLine("      explanation: \"\"");
                            sb.AppendLine($"      isCorrect: {line.Contains("[x]")}");
                        }
                    }

                    lines.RemoveRange(i, lastLine - i);
                    break;
                }
            }

            return sb.ToString();
        }
    }
}