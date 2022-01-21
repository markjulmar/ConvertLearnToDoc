using System.Drawing;
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
            currentParagraph.Style("CodeBlock");
            currentParagraph.Append(new Paragraph(language).Style("CodeFooter"));
        }

        private static void AddBlockedCodeStyle(IDocument document)
        {
            if (!document.Styles.HasStyle("SourceCodeChar", StyleType.Character))
            {
                var codeStyle = document.Styles.AddStyle("SourceCodeChar", StyleType.Character);
                codeStyle.Name = "Source Code Char";
                codeStyle.Linked = "SourceCodeBlock";
                codeStyle.BasedOn = "DefaultParagraphFont";
                codeStyle.Formatting.Font = Globals.CodeFont;
                codeStyle.Formatting.FontSize = Globals.CodeFontSize;
            }

            if (!document.Styles.HasStyle("CodeBlock", StyleType.Paragraph))
            {
                var codeStyle = document.Styles.AddStyle("CodeBlock", StyleType.Paragraph);
                codeStyle.Name = "Code block";
                codeStyle.BasedOn = "Normal";
                codeStyle.Linked = "SourceCodeChar";
                codeStyle.ParagraphFormatting.WordWrap = false;
                codeStyle.ParagraphFormatting.LineSpacingAfter = 0;
                codeStyle.ParagraphFormatting.SetBorders(BorderStyle.Single, Color.LightGray, 5, 2);
                codeStyle.ParagraphFormatting.ShadeFill = Globals.CodeBoxShade;
                codeStyle.Formatting.Font = Globals.CodeFont;
                codeStyle.Formatting.FontSize = Globals.CodeFontSize;
            }

            if (!document.Styles.HasStyle("CodeFooter", StyleType.Paragraph))
            {
                var codeStyle = document.Styles.AddStyle("CodeFooter", StyleType.Paragraph);
                codeStyle.Name = "Code Footer";
                codeStyle.BasedOn = "Normal";
                codeStyle.Formatting.Italic = true;
                codeStyle.Formatting.Superscript = true;
                codeStyle.ParagraphFormatting.LineSpacingBefore = 0;
                codeStyle.ParagraphFormatting.LineSpacingAfter = 5;
                codeStyle.ParagraphFormatting.Alignment = Alignment.Right;
            }
        }
    }
}