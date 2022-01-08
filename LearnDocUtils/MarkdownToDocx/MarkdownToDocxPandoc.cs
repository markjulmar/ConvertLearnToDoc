using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DXPlus;
using MSLearnRepos;

namespace LearnDocUtils
{
    sealed class MarkdownToDocxPandoc : IMarkdownToDocx
    {
        public async Task Convert(TripleCrownModule moduleData, string markdownFile, string docxFile, string zonePivot)
        {
            // Do some pre-processing on the Markdown file.
            var folder = Path.GetDirectoryName(markdownFile);
            var processedMarkdownFile = Path.Combine(folder, Path.ChangeExtension(Path.GetRandomFileName(), "md"));

            string markdownText = PreprocessMarkdownText(File.ReadAllText(markdownFile));
            using (var tempFile = new StreamWriter(processedMarkdownFile))
            {
                await tempFile.WriteLineAsync("---");
                await tempFile.WriteLineAsync($"title: {moduleData.Title}");
                await tempFile.WriteLineAsync("---");
                await tempFile.WriteLineAsync();
                await tempFile.WriteAsync(markdownText);
            }

            try
            {
                // Convert the file.
                await PandocUtils.ConvertFileAsync(processedMarkdownFile, docxFile, folder,
                    "-f markdown-fenced_divs", "-t docx");

                // Post-process the file.
                PostProcessDocument(moduleData, docxFile);
            }
            finally
            {
                File.Delete(processedMarkdownFile);
            }
        }

        static string PreprocessMarkdownText(string text)
        {
            text = Regex.Replace(text, @"\[!include\[(.*?)\]\((.*?)\)]", m => $"#include \"{m.Groups[2].Value}\"", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<rgn>(.*?)</rgn>", _ => "@@rgn@@", RegexOptions.IgnoreCase);
            text = ConvertTripleColonImagesToTags(text);
            //text = ConvertVideoTags(text);
            //text = ConvertTripleColonTables(text);

            return text;
        }

        static void PostProcessDocument(TripleCrownModule moduleData, string docxFile)
        {
            using var doc = Document.Load(docxFile);

            // Add the metadata
            doc.SetPropertyValue(DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);
            doc.SetPropertyValue(DocumentPropertyName.Subject, moduleData.Summary);

            doc.AddCustomProperty("ModuleUid", moduleData.Uid);
            doc.AddCustomProperty("MsTopic", moduleData.Metadata.MsTopic);
            doc.AddCustomProperty("MsProduct", moduleData.Metadata.MsProduct);
            doc.AddCustomProperty("Abstract", moduleData.Abstract);

            List<Paragraph> captions = new();
            var paragraphs = doc.Paragraphs.ToList();
            for (var index = 0; index < paragraphs.Count; index++)
            {
                var paragraph = paragraphs[index];

                // Go through and add highlights to all custom Markdown extensions.
                foreach (var (_, _) in paragraph.FindPattern(new Regex(":::(.*?):::")))
                {
                    paragraph.Runs.ToList()
                        .ForEach(run => run.AddFormatting(new Formatting { Highlight = Highlight.Yellow }));
                }

                captions.AddRange(from _ in paragraph.Pictures
                                  where paragraph.Runs.Count() == 1
                                  select doc.Paragraphs.ElementAt(index + 1));
            }

            // Remove the captions on pictures.
            captions.ForEach(p => p.SetText(string.Empty));

            doc.Save();
            doc.Close();
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
                int end = text.IndexOf(endText, index + startText.Length, StringComparison.InvariantCultureIgnoreCase);
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
