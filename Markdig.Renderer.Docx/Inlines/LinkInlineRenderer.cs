namespace Markdig.Renderer.Docx.Inlines;

public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkInline link)
    {
        string lightboxImageUrl = null;
        var quoteBlockOwner = link.Parent?.ParentBlock?.Parent as QuoteSectionNoteBlock;

        // Look for a lightbox or URL with image
        // [![{alt-text}](image-url)](image-url#lightbox)
        if (link.FirstOrDefault() is LinkInline li)
        {
            if (li.IsImage && !link.IsImage && link.Url != null)
            {
                // [![alt-text](image.png)](url)
                Write(owner, document, currentParagraph, li);
                var image = currentParagraph.Drawings.FirstOrDefault();
                if (image?.Picture != null)
                {
                    var uri = link.Url.StartsWith("http", StringComparison.InvariantCultureIgnoreCase)
                        ? new Uri(link.Url, UriKind.Absolute)
                        : link.Url.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase)
                            ? new Uri(owner.ConvertRelativeUrl(link.Url[..^3]), UriKind.Absolute)
                            : new Uri(link.Url, UriKind.RelativeOrAbsolute);
                    image.Hyperlink = uri;
                }

                return;
            }

            lightboxImageUrl = li.Url;
        }

        var url = link.GetDynamicUrl?.Invoke() ?? link.Url;
        if (url?.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase) == true)
        {
            // Relative links converted to absolute.
            url = new Uri(owner.ConvertRelativeUrl(url[..^3]), UriKind.Absolute).OriginalString;
        }

        string title = link.Title;
        if (string.IsNullOrEmpty(title))
        {
            if (link.FirstChild is LiteralInline literal)
                title = literal.Content.ToString();
        }

        if (link.IsImage)
        {
            bool addBorder = quoteBlockOwner?.SectionAttributeString != null 
                        && quoteBlockOwner.SectionAttributeString.Contains("mx-imgBorder");

            string description = null;

            // If we have text children, then make that our alt-text
            if (link.Any())
            {
                description = title;
                title = string.Join("\r\n", link.Select(il => il.ToString()));
                if (title == description)
                    description = null;
            }

            owner.InsertImage(currentParagraph, link, url, title, description, addBorder);
            if (!string.IsNullOrEmpty(lightboxImageUrl))
                owner.AddComment(currentParagraph, $"lightbox:\"{lightboxImageUrl}\"");
        }
        else
        {
            if (string.IsNullOrEmpty(title))
                title = url;

            try
            {
                currentParagraph.Add(new Hyperlink(title??"", new Uri(url??"", UriKind.RelativeOrAbsolute)));
            }
            catch
            {
                currentParagraph.AddText($"{title} ({url})");
            }
        }
    }
}