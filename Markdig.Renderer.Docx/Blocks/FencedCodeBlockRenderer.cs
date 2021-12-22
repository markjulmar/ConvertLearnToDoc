using System;
using System.Drawing;
using System.Linq;
using DXPlus;
using Markdig.Syntax;

namespace Markdig.Renderer.Docx.Blocks
{
    public class FencedCodeBlockRenderer : DocxObjectRenderer<FencedCodeBlock>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, FencedCodeBlock fencedCodeBlock)
        {
            AddBlockedCodeStyle(document);

            string language = fencedCodeBlock?.Info;

            currentParagraph ??= document.AddParagraph();

            WriteChildren(fencedCodeBlock, owner, document, currentParagraph);
            currentParagraph.Style("SourceCodeBlock")
                .AttachComment(document.CreateComment(Environment.UserName, language));
        }

        private void AddBlockedCodeStyle(IDocument document)
        {
            if (!document.Styles.HasStyle("SourceCodeBlock", StyleType.Paragraph))
            {
                var codeStyle = document.Styles.AddStyle("SourceCodeBlock", StyleType.Paragraph);
                codeStyle.BasedOn = "SourceCode";
                codeStyle.ParagraphFormatting.SetBorders(BorderStyle.Single, Color.LightGray, 5, 2);
                codeStyle.ParagraphFormatting.ShadeFill = Color.FromArgb(0xf0, 0xf0, 0xf0);
            }
        }
    }
}