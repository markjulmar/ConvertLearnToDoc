using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;

namespace LearnDocUtils
{
    public sealed class DocxToLearnPandoc : IDocxToLearn
    {
        private Action<string> logger;

        public async Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder, Action<string> logger)
        {
            this.logger = logger ?? Console.WriteLine;

            if (string.IsNullOrEmpty(mediaFolder))
                mediaFolder = Path.GetDirectoryName(markdownFile);
            else if (!Path.IsPathRooted(mediaFolder))
                mediaFolder = Path.Combine(Path.GetDirectoryName(markdownFile), mediaFolder);
            if (!Directory.Exists(mediaFolder))
                Directory.CreateDirectory(mediaFolder);

            try
            {
                await Utils.ConvertFileAsync(this.logger, docxFile, markdownFile, outputFolder,
                    $"--extract-media=\"{mediaFolder}\"", "--wrap=none", "-t markdown-simple_tables-multiline_tables-grid_tables+pipe_tables");

                var moduleBuilder = new ModuleBuilder(docxFile, outputFolder, tempFile, logger);
                await moduleBuilder.CreateModuleAsync(ProcessMarkdown);
            }
            finally
            {
                if (!keepTempFiles)
                {
                    File.Delete(tempFile);
                }
                else
                {
                    this.logger?.Invoke($"Consolidated Markdown file: {tempFile}");
                }
            }
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
