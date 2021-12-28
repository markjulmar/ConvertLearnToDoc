using System;
using System.IO;
using DXPlus;

namespace Docx.Renderer.Markdown.Renderers
{
    public class RunRenderer : MarkdownObjectRenderer<Run>
    {
        protected override void Render(IMarkdownRenderer renderer, TextWriter writer, Run element, object tags)
        {
            dynamic md = tags;
            TextFormatting tf = md.TextFormatting;
            if (element.Properties.Bold)
                tf.Bold = element.Properties.Bold;
            if (element.Properties.Italic)
                tf.Italic = element.Properties.Italic;
            if (!string.IsNullOrEmpty(tf.StyleName))
                tf.StyleName = element.StyleName;

            string surrounds =
                tf.Bold ? "**"
                : tf.Italic ? "*" : null;

            foreach (var e in element.Elements)
            {
                if (e is Text t)
                {
                    if (surrounds != null) writer.Write(surrounds);
                    writer.Write(t.ToString());
                    if (surrounds != null) writer.Write(surrounds);
                }
                else if (e is Break b)
                {
                    if (b.Type == BreakType.Line)
                        writer.Write(Environment.NewLine);
                }
                else if (e is Drawing d)
                {
                }
            }
        }
    }
}
