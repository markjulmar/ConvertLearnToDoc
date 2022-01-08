using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Linq;

namespace LearnDocUtils
{
    sealed class DocxToMarkdownPandoc : IDocxToMarkdown
    {
        public async Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder)
        {
            if (docxFile == null) 
                throw new ArgumentNullException(nameof(docxFile));
            if (string.IsNullOrEmpty(markdownFile)) 
                throw new ArgumentNullException(nameof(markdownFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"{docxFile} does not exist.", nameof(docxFile));
            if (Path.GetInvalidFileNameChars().Any(markdownFile.Contains))
                throw new ArgumentException($"{markdownFile} is an invalid filename.", nameof(markdownFile));

            if (File.Exists(markdownFile))
                File.Delete(markdownFile);

            if (string.IsNullOrEmpty(mediaFolder))
                mediaFolder = Path.Combine(Path.GetDirectoryName(markdownFile)??"", "media");
            if (!Directory.Exists(mediaFolder))
                Directory.CreateDirectory(mediaFolder);

            string outputFolder = Path.GetDirectoryName(markdownFile);
            outputFolder = string.IsNullOrEmpty(outputFolder) ? Directory.GetCurrentDirectory() : Path.GetFullPath(outputFolder);

            await PandocUtils.ConvertFileAsync(docxFile, markdownFile, outputFolder,
                $"--extract-media=\"{mediaFolder}\"", "--wrap=none", "-t markdown-simple_tables-multiline_tables-grid_tables+pipe_tables");

            // Do some post-processing.
            string markdownText = PostProcessMarkdown(File.ReadAllText(markdownFile));
            File.WriteAllText(markdownFile, markdownText);
        }

        private static string PostProcessMarkdown(string text)
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
