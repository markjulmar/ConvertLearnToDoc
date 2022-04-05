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
        if (!document.Styles.Exists("SourceCodeChar", StyleType.Character))
        {
            var codeStyle = document.Styles.Add("SourceCodeChar", "Source Code Char", StyleType.Character);
            codeStyle.LinkedStyle = "SourceCodeBlock";
            codeStyle.BasedOn = "DefaultParagraphFont";
            codeStyle.Formatting = new()
            {
                Font = Globals.CodeFont,
                FontSize = Globals.CodeFontSize
            };
        }

        if (!document.Styles.Exists("CodeBlock", StyleType.Paragraph))
        {
            var codeStyle = document.Styles.Add("CodeBlock", "Code Block", StyleType.Paragraph);
            codeStyle.BasedOn = "Normal";
            codeStyle.LinkedStyle = "SourceCodeChar";
            codeStyle.ParagraphFormatting = new()
            {
                WordWrap = false,
                LineSpacingAfter = 0,
                Shading = new Shading { Fill = Globals.CodeBoxShade }
            };
            codeStyle.ParagraphFormatting.SetOutsideBorders(new Border { Style = BorderStyle.Single, Color = Color.LightGray, Spacing = 5, Size = 2 });
            codeStyle.Formatting = new()
            {
                Font = Globals.CodeFont,
                FontSize = Globals.CodeFontSize
            };
        }

        if (!document.Styles.Exists("CodeFooter", StyleType.Paragraph))
        {
            var codeStyle = document.Styles.Add("CodeFooter", "Code Footer", StyleType.Paragraph);
            codeStyle.BasedOn = "Normal";
            codeStyle.Formatting = new()
            {
                Italic = true,
                Superscript = true
            };
            codeStyle.ParagraphFormatting = new()
            {
                LineSpacingBefore = 0,
                LineSpacingAfter = 5,
                Alignment = Alignment.Right
            };
        }
    }
}