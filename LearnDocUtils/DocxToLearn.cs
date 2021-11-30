using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DXPlus;
using System.Reflection;
using System.IO;

namespace LearnDocUtils
{
    public static class DocxToLearn
    {
        public static async Task ConvertAsync(string docxFile, string outputFolder)
        {
            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or whitespace.", nameof(outputFolder));
            }

            if (string.IsNullOrWhiteSpace(docxFile))
            {
                throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
            }

            if (!File.Exists(docxFile))
            {
                Console.WriteLine($"Error: {docxFile} does not exist.");
                return;
            }

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Convert to Markdown.
            string tempFile = Path.Combine(outputFolder, "temp.md");
            await Utils.ConvertFileAsync(docxFile, tempFile, outputFolder, 
                "--extract-media=.", "--wrap=none", "-t markdown-simple_tables-multiline_tables-grid_tables+pipe_tables");

            // Now pick off title metadata.
            var metadata = LoadDocumentMetadata(docxFile);

            string includeFolder = Path.Combine(outputFolder, "includes");
            Directory.CreateDirectory(includeFolder);

            // Read all the lines
            var files = new Dictionary<string, List<string>>();
            using (var reader = new StreamReader(tempFile))
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
                string unitFileName = index.ToString() + "-" + baseFn;

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
                await File.WriteAllTextAsync(Path.Combine(includeFolder, Path.ChangeExtension(unitFileName, "md")), ProcessMarkdown(string.Join("\r\n", value)));
                unitIds.Add($"- {moduleUid}.{baseFn}");
                index++;
            }

            // Write the index.yml file.
            var moduleValues = new Dictionary<string, string>()
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

            File.Delete(tempFile);
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
                    sb.AppendLine($"  title: {quiz[0].Substring(3)}");
                    sb.AppendLine("  questions:");

                    for (int pos = 1; pos < quiz.Count; pos++)
                    {
                        string line = quiz[pos];
                        if (string.IsNullOrWhiteSpace(line))
                            continue;
                        if (line.StartsWith("### "))
                        {
                            sb.AppendLine($"  - content: \"{line.Substring(4)}\"");
                            sb.AppendLine( "    choices:");
                        }
                        else
                        {
                            sb.AppendLine($"    - content: \"{line[(line.IndexOf(']') + 1)..].TrimStart()}\"");
                            sb.AppendLine( "      explanation: \"\"");
                            sb.AppendLine($"      isCorrect: {line.Contains("[x]")}");
                        }
                    }

                    lines.RemoveRange(i, lastLine - i);
                    break;
                }
            }

            return sb.ToString();
        }

        private static string ProcessMarkdown(string text)
        {
            text = text.Trim('\r').Trim('\n');

            text = Regex.Replace(text, @" \\\[!TIP\\\] ", " [!TIP] ");
            text = Regex.Replace(text, @" \\\[!NOTE\\\] ", " [!NOTE] ");
            text = Regex.Replace(text, @" \\\[!WARNING\\\] ", " [!WARNING] ");
            text = Regex.Replace(text, @"@@rgn@@", "<rgn>[sandbox resource group name]</rgn>");
            text = Regex.Replace(text, @"{width=""[0-9.]*in""\s+height=""[0-9.]*in""}\s*", "\r\n");
            text = Regex.Replace(text, @"#include ""(.*?)""", m => $"[!include []({m.Groups[1].Value})]");
            text = text.Replace("(./media/", "(../media/");
            text = text.Replace((char)0xa0, ' ');

            return text;
        }
    }
}
