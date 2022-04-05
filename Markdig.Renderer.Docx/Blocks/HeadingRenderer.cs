using System.Drawing;

namespace Markdig.Renderer.Docx.Blocks;

public class HeadingRenderer : DocxObjectRenderer<HeadingBlock>
{
    private const string TabConceptualMarker = "#tab/";

    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, HeadingBlock heading)
    {
        // Allowed in a row extension.
        currentParagraph ??= document.AddParagraph();

        // Test for tab-conceptual link.
        if (heading.Inline?.FirstChild is LinkInline link)
        {
            if (link.Url?.StartsWith(TabConceptualMarker) == true)
            {
                ProcessTabConceptualHeader(currentParagraph, link);
                return;
            }
        }

        switch (heading.Level)
        {
            case 1: currentParagraph.Style(HeadingType.Heading1); break;
            case 2: currentParagraph.Style(HeadingType.Heading2); break;
            case 3: currentParagraph.Style(HeadingType.Heading3); break;
            case 4: currentParagraph.Style(HeadingType.Heading4); break;
            case 5: currentParagraph.Style(HeadingType.Heading5); break;
        }

        WriteChildren(heading, owner, document, currentParagraph);
    }

    private static void ProcessTabConceptualHeader(Paragraph currentParagraph, LinkInline tabGroup)
    {
        string url = tabGroup.Url ?? TabConceptualMarker;
        string title = tabGroup.Title;
        if (string.IsNullOrEmpty(title))
            title = (tabGroup.FirstOrDefault() as LiteralInline)?.Content.ToString() ?? "";

        currentParagraph.AddText(new Run($"{{tabgroup \"{title}\" {url[TabConceptualMarker.Length..]}}}",
            new Formatting { Highlight = Highlight.DarkGray, Color = Color.White }));
    }
}