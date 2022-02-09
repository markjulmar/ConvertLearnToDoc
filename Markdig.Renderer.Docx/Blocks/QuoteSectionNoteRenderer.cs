using System;
using DXPlus;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.Blocks
{
    public class QuoteSectionNoteRenderer : DocxObjectRenderer<QuoteSectionNoteBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, QuoteSectionNoteBlock block)
        {
            currentParagraph ??= document.AddParagraph();

            if (block.QuoteType == QuoteSectionNoteType.DFMNote 
                && block.NoteTypeString != null)
            {
                // If there was a note block right before this one, add a separator
                // otherwise Word merges them together.
                if (currentParagraph.PreviousParagraph.Properties.StyleName == HeadingType.IntenseQuote.ToString())
                    currentParagraph.InsertBefore(new Paragraph());

                currentParagraph.Style(HeadingType.IntenseQuote).AppendLine(block.NoteTypeString);
            }
            else if (block.QuoteType == QuoteSectionNoteType.DFMVideo)
            {
                string videoLink = block.VideoLink;

                using var placeholder = owner.GetEmbeddedResource("video-placeholder.png");
                currentParagraph.Properties.Alignment = Alignment.Center;
                currentParagraph.Append(document.CreateVideo(
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
                    currentParagraph.AppendLine()
                            .AppendLine();
                }

                Write(paragraphBlock, owner, document, currentParagraph);
            }
        }
    }
}
