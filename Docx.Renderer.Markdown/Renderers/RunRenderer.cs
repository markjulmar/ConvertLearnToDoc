using System;
using System.IO;
using DXPlus;
using Julmar.GenMarkdown;
using DXText = DXPlus.Text;
using Text = Julmar.GenMarkdown.Text;

namespace Docx.Renderer.Markdown.Renderers
{
    public class RunRenderer : MarkdownObjectRenderer<Run>
    {
        protected override void Render(IMarkdownRenderer renderer, 
            MarkdownDocument document, MarkdownBlock blockOwner, 
            Run element, RenderBag tags)
        {
            if (blockOwner == null)
            {
                blockOwner = new Julmar.GenMarkdown.Paragraph();
                document.Add(blockOwner);
            }

           var tf = tags.Get<TextFormatting>(nameof(TextFormatting));
            if (element.Properties.Bold)
                tf.Bold = element.Properties.Bold;
            if (element.Properties.Italic)
                tf.Italic = element.Properties.Italic;
            if (!string.IsNullOrEmpty(tf.StyleName))
                tf.StyleName = element.StyleName;

            foreach (var e in element.Elements)
            {
                switch (e)
                {
                    case DXText t:
                    {
                        var p = (Julmar.GenMarkdown.Paragraph) blockOwner;

                        if (element.Parent is Hyperlink hl)
                        {
                            p.Add(Text.Link(hl.Text, hl.Uri.ToString()));
                        }
                        else
                        {
                            if (tf.Bold && tf.Italic)
                                p.Add($"**_{t.Value}_**");
                            else if (tf.Bold)
                                p.Add(Text.Bold(t.Value));
                            else if (tf.Italic)
                                p.Add(Text.Italic(t.Value));
                            else if (tf.Monospace)
                                p.Add(Text.Code(t.Value));
                            else
                                p.Add(t.Value);
                        }
                        break;
                    }
                    case Break b:
                    {
                        var p = (Julmar.GenMarkdown.Paragraph)blockOwner;
                        if (b.Type == BreakType.Line)
                            p.Add(Text.LineBreak);
                        break;
                    }
                    case Drawing {Picture: {}} d:
                    {
                        var p = d.Picture;

                        string fn = Path.Combine(renderer.MediaFolder, p.FileName);
                        using var input = p.Image.OpenStream();
                        using var output = File.OpenWrite(fn);
                        input.CopyTo(output);

                        document.Remove(blockOwner);
                        document.Add(new Julmar.GenMarkdown.Image(p.Description, fn));
                        break;
                    }
                }
            }
        }
    }
}
