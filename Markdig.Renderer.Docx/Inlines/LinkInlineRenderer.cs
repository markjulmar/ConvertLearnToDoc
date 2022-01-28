using System;
using System.Linq;
using DXPlus;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Inlines
{
    public class LinkInlineRenderer : DocxObjectRenderer<LinkInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkInline link)
        {
            bool isLightboxImage = false;
            var quoteBlockOwner = link.Parent?.ParentBlock?.Parent as QuoteSectionNoteBlock;

            // Look for a lightbox image
            // [![{alt-text}](image-url)](image-url#lightbox)
            if (link.FirstOrDefault() is LinkInline li)
            {
                // Render that as the image
                link = li;
                isLightboxImage = true;
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
                bool addBorder = false;
                if (quoteBlockOwner != null)
                {
                    if (quoteBlockOwner.SectionAttributeString != null
                        && quoteBlockOwner.SectionAttributeString.Contains("mx-imgBorder"))
                        addBorder = true;
                }

                owner.InsertImage(currentParagraph, url, title, addBorder, isLightboxImage);
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
}