﻿using System;
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
                currentParagraph
                    .Style(HeadingType.IntenseQuote)
                    .AppendLine(block.NoteTypeString);
            }
            else if (block.QuoteType == QuoteSectionNoteType.DFMVideo)
            {
                string videoLink = block.VideoLink;
                // currentParagraph.Append($"{{video: {videoLink}}}",
                //     new Formatting { Highlight = Highlight.Magenta, Color = Color.White });

                using var placeholder = owner.GetEmbeddedResource("video-placeholder.png");
                currentParagraph.Properties.Alignment = Alignment.Center;
                currentParagraph.Append(document.CreateVideo(
                    placeholder, ImageContentType.Png,
                    new Uri(videoLink, UriKind.Absolute),
                    400, 225));
                
                return;
            }

            WriteChildren(block, owner, document, currentParagraph);
        }
    }
}
