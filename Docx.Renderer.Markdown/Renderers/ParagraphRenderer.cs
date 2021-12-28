using System.IO;
using DXPlus;

namespace Docx.Renderer.Markdown.Renderers
{
    public sealed class ParagraphRenderer : MarkdownObjectRenderer<Paragraph>
    {
        protected override void Render(IMarkdownRenderer renderer, TextWriter writer, Paragraph element, object tags)
        {
            TextFormatting tf = new()
            {
                StyleName = element.Properties.StyleName,
                Bold = element.Properties.DefaultFormatting.Bold,
                Italic = element.Properties.DefaultFormatting.Italic
            };

            WritePrefix(writer, tf.StyleName);
            renderer.WriteContainer(writer, element.Runs, new { TextFormatting = tf, Paragraph = element });
        }

        private static void WritePrefix(TextWriter writer, string styleName)
        {
            string prefix = styleName ?? "" switch
            {
                "Heading1" => "#",
                "Heading2" => "##",
                "Heading3" => "###",
                "Heading4" => "####",
                "Heading5" => "#####",
                _ => string.Empty
            };
            if (!string.IsNullOrEmpty(prefix))
                writer.Write(prefix + " ");
        }
    }
}
