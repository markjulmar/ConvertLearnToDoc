using System;
using System.IO;
using Docx.Renderer.Markdown;

namespace CreateDoc
{
    internal class Program
    {
        static void Main()
        {
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string docxFile = Path.Combine(desktopFolder, "test.docx");
            string markdownFile = Path.Combine(desktopFolder, "out.md");
            string mediaFolder = Path.Combine(desktopFolder, "media");

            new MarkdownRenderer().Convert(docxFile, markdownFile, mediaFolder);
            Console.WriteLine(File.ReadAllText(markdownFile));
        }
    }
}
