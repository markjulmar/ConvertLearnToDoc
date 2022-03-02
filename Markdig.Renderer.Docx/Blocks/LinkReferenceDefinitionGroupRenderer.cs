namespace Markdig.Renderer.Docx.Blocks
{
    public class LinkReferenceDefinitionGroupRenderer : DocxObjectRenderer<LinkReferenceDefinitionGroup>
    {
        public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, LinkReferenceDefinitionGroup group)
        {
            if (group.Links.Count > 0)
            {

            }
        }
    }
}
