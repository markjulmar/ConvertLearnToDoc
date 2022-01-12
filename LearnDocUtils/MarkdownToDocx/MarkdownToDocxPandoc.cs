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
            var folder = Path.GetDirectoryName(markdownFile)??"";
            var processedMarkdownFile = Path.Combine(folder, Path.ChangeExtension(Path.GetRandomFileName(), "md"));

            string markdownText = PreprocessMarkdownText(await File.ReadAllTextAsync(markdownFile));
            await File.WriteAllTextAsync(processedMarkdownFile, markdownText);

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

        private static void PostProcessDocument(TripleCrownModule moduleData, string docxFile)
        {
            using var doc = Document.Load(docxFile);

            // Add the metadata
            WriteTitle(moduleData, doc);
            AddMetadata(moduleData, doc);

            var paragraphs = doc.Paragraphs.ToList();
            foreach (var paragraph in paragraphs)
            {
                // Go through and add highlights to all custom Markdown extensions.
                foreach (var (_, _) in paragraph.FindPattern(new Regex(":::(.*?):::")))
                {
                    paragraph.Runs.ToList()
                        .ForEach(run => run.AddFormatting(new Formatting { Highlight = Highlight.Yellow }));
                }
            }

            // Remove the captions on pictures.
            paragraphs
                .Where(p => p.Properties.StyleName == "ImageCaption")
                .ToList()
                .ForEach(p => p.Remove());

            doc.Save();
            doc.Close();
        }

        private static void WriteTitle(TripleCrownModule moduleData, IDocument document)
        {
            var firstParagraph = document.Paragraphs.First();
            firstParagraph
                .InsertBefore(new Paragraph(
                            $"Last modified on {moduleData.LastUpdated.ToShortDateString()} by {moduleData.Metadata.MsAuthor}@microsoft.com")
                        .Style(HeadingType.Subtitle))
                .InsertBefore(new Paragraph(moduleData.Title)
                        .Style(HeadingType.Title));
        }

        private static void AddMetadata(TripleCrownModule moduleData, IDocument document)
        {
            document.SetPropertyValue(DocumentPropertyName.Title, moduleData.Title);
            document.SetPropertyValue(DocumentPropertyName.Subject, moduleData.Summary);
            document.SetPropertyValue(DocumentPropertyName.Keywords, string.Join(',', moduleData.Products));
            document.SetPropertyValue(DocumentPropertyName.Comments, string.Join(',', moduleData.FriendlyLevels));
            document.SetPropertyValue(DocumentPropertyName.Category, string.Join(',', moduleData.FriendlyRoles));
            document.SetPropertyValue(DocumentPropertyName.CreatedDate, moduleData.LastUpdated.ToString("yyyy-MM-ddT00:00:00Z"));
            document.SetPropertyValue(DocumentPropertyName.Creator, moduleData.Metadata.MsAuthor);

            // Add custom data.
            document.AddCustomProperty("ModuleUid", moduleData.Uid);
            document.AddCustomProperty("MsTopic", moduleData.Metadata.MsTopic);
            document.AddCustomProperty("MsProduct", moduleData.Metadata.MsProduct);
            document.AddCustomProperty("Abstract", moduleData.Abstract);
        }
        
        /* Pandoc can't handle HTML tables.
         * Removing method.
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
        */

        /* Pandoc can't handle video tags either :(
         *
        private static string ConvertVideoTags(string text)
        {
            var matches = Regex.Matches(text, @">[ ]*\[!VIDEO (.*?)\]", RegexOptions.IgnoreCase);
            foreach (Match m in matches)
            {
                string url = m.Groups[1].Value;
                text = text.Replace(m.Value,
                    $"<video width=\"640\" height=\"480\" controls>\r\n\t<source src=\"{url}\" type=\"video/mp4\">\r\n</video>\r\n");
            }
            return text;
        }
        */

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
            string result = match.Groups[1].Value;
            return string.IsNullOrEmpty(result) ? null : result;
        }
    }
}
