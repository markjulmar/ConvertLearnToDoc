namespace Markdig.Renderer.Docx.Inlines;

public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkInline link)
    {
        string lightboxImageUrl = null;
        var quoteBlockOwner = link.Parent?.ParentBlock?.Parent as QuoteSectionNoteBlock;

        // Look for a lightbox image
        // [![{alt-text}](image-url)](image-url#lightbox)
        if (link.FirstOrDefault() is LinkInline li)
        {
            lightboxImageUrl = li.Url;
        }

        var url = link.GetDynamicUrl?.Invoke() ?? link.Url;

        string title = link.Title;
        if (string.IsNullOrEmpty(title))
        {
            if (link.FirstChild is LiteralInline literal)
                title = literal.Content.ToString();
        }

        if (link.IsImage)
        {
            bool addBorder = quoteBlockOwner?.SectionAttributeString != null && quoteBlockOwner.SectionAttributeString.Contains("mx-imgBorder");

            string description = null;

            // If we have text children, then make that our alt-text
            if (link.Any())
            {
                description = title;
                title = string.Join("\r\n", link.Select(il => il.ToString()));
                if (title == description)
                    description = null;
            }

            owner.InsertImage(currentParagraph, link, url, title, description, addBorder, lightboxImageUrl);
        }
        else
        {
            if (string.IsNullOrEmpty(title))
                title = url;

            try
            {
                currentParagraph.Append(new Hyperlink(title, new Uri(url??"", UriKind.RelativeOrAbsolute)));
            }
            catch
            {
                currentParagraph.Append($"{title} ({url})");
            }
        }
    }
}