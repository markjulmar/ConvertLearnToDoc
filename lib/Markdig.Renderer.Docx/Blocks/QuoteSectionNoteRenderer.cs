using System.Globalization;

namespace Markdig.Renderer.Docx.Blocks;

public class QuoteSectionNoteRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteSectionNoteBlock block)
    {
        currentParagraph ??= document.AddParagraph();

        if (block.QuoteType is QuoteSectionNoteType.DFMNote or QuoteSectionNoteType.MarkdownQuote)
        {
            var style = block.QuoteType == QuoteSectionNoteType.DFMNote
                ? HeadingType.IntenseQuote
                : HeadingType.Quote;

            // If there was a note block right before this one, add a separator
            // otherwise Word merges them together.
            if (currentParagraph.PreviousParagraph?.Properties.StyleName == style.ToString())
                currentParagraph.InsertBefore(new Paragraph());

            if (!string.IsNullOrEmpty(block.NoteTypeString))
                currentParagraph.Style(style).AddRange(new [] {
                    new Run(CultureInfo.CurrentCulture.TextInfo.ToTitleCase(block.NoteTypeString.ToLower()) + ":", 
                        new Formatting { Bold = true, Underline = true }), 
                    new Run(" ")
                });
            else
                currentParagraph.Style(style);
        }
        else if (block.QuoteType == QuoteSectionNoteType.DFMVideo)
        {
            string videoLink = block.VideoLink;

            using var placeholder = owner.GetEmbeddedResource("video-placeholder.png");
            currentParagraph.Properties.Alignment = Alignment.Center;
            currentParagraph.Add(document.CreateVideo(
                placeholder, ImageContentType.Png,
                new Uri(videoLink, UriKind.Absolute),
                400, 225));
                
            return;
        }

        // Write all the paragraphs with newlines between.
        for (var index = 0; index < block.Count; index++)
        {
            var paragraphBlock = block[index];
            if (index > 0)
            {
                currentParagraph
                    .Newline()
                    .Newline();
            }

            Write(paragraphBlock, owner, document, currentParagraph);
        }
    }
}