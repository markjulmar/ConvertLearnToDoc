using System.Diagnostics;

namespace Markdig.Renderer.Docx.Inlines;

public class EmphasisInlineRenderer : DocxObjectRenderer<EmphasisInline>
{
    public override void Write(IDocxRenderer owner, IDocument document, Paragraph currentParagraph, EmphasisInline emphasis)
    {
        Debug.Assert(currentParagraph != null);

        // If there's more than one literal, then combine them together so we have one
        // string to emphasize in Word. This optimizes the runs so we only have one run of text
        // with the same properties. We also want to ignore line breaks.
        if (emphasis.Count() > 1 && emphasis.All(c => c is LiteralInline || c is LineBreakInline { IsHard:false }))
        {
            string text = string.Join("", emphasis.Select(c => c is LineBreakInline ? " " : ((LiteralInline) c).Content.ToString()));
            emphasis.Clear();
            emphasis.AppendChild(new LiteralInline(text));
        }

        // Write children into the paragraph..
        WriteChildren(emphasis, owner, document, currentParagraph);
        // .. and then change the style of that run.
        if (emphasis.DelimiterChar is '*' or '_')
        {
            currentParagraph.WithFormatting(emphasis.DelimiterCount == 2
                ? new Formatting {Bold = true}
                : new Formatting {Italic = true});
        }
    }
}