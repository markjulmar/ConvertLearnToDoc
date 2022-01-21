using DXPlus;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderer.Docx.Inlines
{
    public class CodeInlineRenderer : DocxObjectRenderer<CodeInline>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, CodeInline obj)
        {
            string code = obj.Content;
            currentParagraph
                .Append(code)
                .WithFormatting(new Formatting
                {
                    Font = Globals.CodeFont, 
                    ShadeFill = Globals.CodeBoxShade
                });
        }
    }
}