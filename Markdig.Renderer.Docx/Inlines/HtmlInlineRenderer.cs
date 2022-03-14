using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Markdig.Renderer.Docx.Inlines;

public class HtmlInlineRenderer : DocxObjectRenderer<HtmlInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HtmlInline html)
    {
        Debug.Assert(currentParagraph != null);

        string tag = Helpers.GetTag(html.Tag);
        bool isClose = html.Tag.StartsWith("</");
        switch (tag)
        {
            case "kbd":
                if (isClose)
                {
                    var r = currentParagraph.Runs.Last();
                    r.MergeFormatting(new Formatting {
                        Bold = true, CapsStyle = CapsStyle.SmallCaps, Font = FontFamily.GenericMonospace,
                        Color = Color.Black,
                        Shading = new Shading { Fill = Globals.CodeBoxShade }
                    });
                }

                break;
            case "b":
                if (!isClose)
                    currentParagraph.Add(new Run(Helpers.ReadLiteralTextAfterTag(owner, html), new Formatting { Bold = true }));
                break;
            case "i":
                if (!isClose)
                    currentParagraph.Add(new Run(Helpers.ReadLiteralTextAfterTag(owner, html), new Formatting { Italic = true }));
                break;
            case "a":
                if (!isClose)
                    ProcessRawAnchor(html, owner, document, currentParagraph);
                break;
            case "br":
                if (html.Parent?.ParentBlock is not HeadingBlock)
                    currentParagraph.Newline();
                break;
            case "sup":
                if (!isClose)
                    currentParagraph.Add(new Run(Helpers.ReadLiteralTextAfterTag(owner, html), new Formatting {Superscript = true}));
                break;
            case "sub":
                if (!isClose)
                    currentParagraph.Add(new Run(Helpers.ReadLiteralTextAfterTag(owner, html), new Formatting { Subscript = true }));
                break;
            case "rgn":
                if (!isClose)
                    currentParagraph.Add(new Run($"{{rgn {Helpers.ReadLiteralTextAfterTag(owner, html)}}}", new Formatting { Highlight = Highlight.Cyan }));
                break;
            default:
                currentParagraph.Add(html.Tag);
                Console.WriteLine($"Encountered unsupported HTML tag: {tag}");
                break;
        }
    }

    private static void ProcessRawAnchor(HtmlInline html, IDocxRenderer owner, IDocument document, Paragraph currentParagraph)
    {
        string text = Helpers.ReadLiteralTextAfterTag(owner, html);
        Regex re = new Regex(@"(?inx)
                <a \s [^>]*
                    href \s* = \s*
                        (?<q> ['""] )
                            (?<url> [^""]+ )
                        \k<q>
                [^>]* >");

        // Ignore if we can't find a URL.
        var m = re.Match(html.Tag);
        if (m.Groups.ContainsKey("url") == false)
        {
            if (text.Length > 0)
                currentParagraph.Add(text);
        }
        else
        {
            currentParagraph.Add(new Hyperlink(text, new Uri(m.Groups["url"].Value)));
        }
    }
}