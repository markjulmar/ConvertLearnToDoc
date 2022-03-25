namespace Markdig.Renderer.Docx.Inlines;

public class CodeInlineRenderer : DocxObjectRenderer<CodeInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, CodeInline obj)
    {
        string code = obj.Content;
        currentParagraph.AddText(new Run(code, new Formatting {
                Font = Globals.CodeFont, 
                Shading = new Shading { Fill = Globals.CodeBoxShade }
            }));
    }
}