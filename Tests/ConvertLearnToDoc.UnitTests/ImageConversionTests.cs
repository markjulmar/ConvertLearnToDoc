using System.IO.Compression;
using System.Xml.Linq;
using HtmlAgilityPack;
using Julmar.DocsToMarkdown;
using LearnDocUtils;

namespace ConvertLearnToDoc.UnitTests;

public class ImageConversionTests
{
    [Theory]
    [InlineData("role=\"presentation\"")]
    [InlineData("role=\"presentation\" alt=\"\"")]
    [InlineData("role=\"PRESENTATION\"")]
    public async Task PresentationIconsInTableRenderAsImagesWithExtensionComments(string attributes)
    {
        var markdown = ConvertHtml($"""
            <table>
            <thead><tr><th>License</th><th>Availability</th></tr></thead>
            <tbody><tr>
            <td><img src="images/license/a1-tag.png" {attributes}><strong>Microsoft365 A1 license</strong></td>
            <td><img src="images/license/a1-tag.png" {attributes}> <img src="images/license/su25.png" {attributes}><strong>Student Use license</strong></td>
            </tr></tbody>
            </table>
            """);

        Assert.Contains(":::image type=\"icon\"", markdown);
        Assert.DoesNotContain("type=\"content\"", markdown);

        var folder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(Path.Combine(folder, "images", "license"));
        try
        {
            // A small PNG keeps the test offline while exercising actual Word image insertion.
            var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=");
            await File.WriteAllBytesAsync(Path.Combine(folder, "images", "license", "a1-tag.png"), png);
            await File.WriteAllBytesAsync(Path.Combine(folder, "images", "license", "su25.png"), png);
            var markdownFile = Path.Combine(folder, "license.md");
            var docxFile = Path.Combine(folder, "license.docx");
            await File.WriteAllTextAsync(markdownFile, markdown);

            var errors = await MarkdownToDocConverter.ConvertMarkdownToDocx(
                null, "https://learn.microsoft.com/en-us/microsoft-365/education/guide/0-start/all-license",
                folder, null, markdownFile, docxFile, null, false);

            Assert.Empty(errors);
            using var archive = ZipFile.OpenRead(docxFile);
            using var documentStream = archive.GetEntry("word/document.xml")!.Open();
            var document = XDocument.Load(documentStream);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var cells = document.Descendants(w + "tc").Skip(2).ToList();
            Assert.Equal(2, cells.Count);
            Assert.Single(cells[0].Descendants(w + "drawing"));
            Assert.Equal(2, cells[1].Descendants(w + "drawing").Count());
            var text = string.Concat(document.Descendants(w + "t").Select(t => t.Value));
            Assert.Contains("Microsoft365 A1 license", text);
            Assert.Contains("Student Use license", text);
            Assert.DoesNotContain(":::image", text);
            Assert.All(cells, cell => Assert.Contains(cell.Descendants(w + "r"), run =>
                run.Element(w + "rPr")?.Element(w + "b") != null));

            using var commentsStream = archive.GetEntry("word/comments.xml")!.Open();
            var comments = XDocument.Load(commentsStream);
            Assert.Equal(3, comments.Descendants(w + "comment").Count());
            Assert.All(comments.Descendants(w + "comment"), comment =>
                Assert.Contains("useExtension", comment.Value));
        }
        finally
        {
            Directory.Delete(folder, true);
        }
    }

    [Theory]
    [InlineData("<img src=\"images/license.png\" alt=\"License options\">", true,
        ":::image type=\"content\" source=\"images/license.png\" alt-text=\"License options\":::")]
    [InlineData("<img src=\"images/license.png\" role=\"presentation\">", false,
        "![](images/license.png)")]
    [InlineData("<img src=\"images/license.png\" role=\"presentation\" aria-describedby=\"description\" alt=\"License options\">", true,
        ":::image type=\"complex\" source=\"images/license.png\" alt-text=\"License options\":::")]
    public void ImageConversionPreservesOtherFormats(string html, bool useExtensions, string expected)
    {
        Assert.Equal(expected, ConvertHtml(html, useExtensions));
    }

    private static string ConvertHtml(string html, bool useExtensions = true)
    {
        var document = new HtmlDocument();
        document.LoadHtml(html);
        var converter = new HtmlConverter(
            new Uri("https://learn.microsoft.com/en-us/microsoft-365/education/guide/0-start/all-license"),
            document, new HtmlConverterOptions { UseDocfxExtensions = useExtensions });
        return converter.Convert(document.DocumentNode.FirstChild);
    }
}
