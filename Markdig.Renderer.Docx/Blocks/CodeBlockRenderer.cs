using System.Drawing;

namespace Markdig.Renderer.Docx.Blocks;

public class CodeBlockRenderer : DocxObjectRenderer<CodeBlock>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, CodeBlock codeBlock)
    {
        AddBlockedCodeStyle(document);
        currentParagraph ??= document.AddParagraph();
        WriteChildren(codeBlock, owner, document, currentParagraph);

        currentParagraph.Style("CodeBlock");

        if (codeBlock is FencedCodeBlock fencedCodeBlock)
        {
            string language = fencedCodeBlock.Info;
            if (!string.IsNullOrEmpty(language))
                currentParagraph.InsertAfter(new Paragraph(language).Style("CodeFooter"));
        }
    }

    private static void AddBlockedCodeStyle(IDocument document)
    {
        if (!document.Styles.HasStyle("SourceCodeChar", StyleType.Character))
        {
            var codeStyle = document.Styles.AddStyle("SourceCodeChar", StyleType.Character);
            codeStyle.Name = "Source Code Char";
            codeStyle.LinkedStyle = "SourceCodeBlock";
            codeStyle.BasedOn = "DefaultParagraphFont";
            codeStyle.Formatting.Font = Globals.CodeFont;
            codeStyle.Formatting.FontSize = Globals.CodeFontSize;
        }

        if (!document.Styles.HasStyle("CodeBlock", StyleType.Paragraph))
        {
            var codeStyle = document.Styles.AddStyle("CodeBlock", StyleType.Paragraph);
            codeStyle.Name = "Code block";
            codeStyle.BasedOn = "Normal";
            codeStyle.LinkedStyle = "SourceCodeChar";
            codeStyle.ParagraphFormatting.WordWrap = false;
            codeStyle.ParagraphFormatting.LineSpacingAfter = 0;
            codeStyle.ParagraphFormatting.SetOutsideBorders(new Border { Style = BorderStyle.Single, Color = Color.LightGray, Spacing = 5, Size = 2 });
            codeStyle.ParagraphFormatting.Shading = new Shading { Fill = Globals.CodeBoxShade };
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