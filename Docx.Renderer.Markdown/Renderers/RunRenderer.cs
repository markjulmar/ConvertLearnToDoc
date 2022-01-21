using System.IO;
using System.Linq;
using DXPlus;
using GenMarkdown.DocFx.Extensions;
using Julmar.GenMarkdown;
using DXText = DXPlus.Text;
using Image = DXPlus.Image;
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
                tf.Bold = true;
            if (element.Properties.Italic)
                tf.Italic = true;
            if (element.Properties.CapsStyle == CapsStyle.SmallCaps)
                tf.KbdTag = true;
            if (!string.IsNullOrEmpty(tf.StyleName))
                tf.StyleName = element.StyleName;
            if (TextFormatting.IsMonospaceFont(element.Properties.Font))
                tf.Monospace = true;

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
                            if (tf.KbdTag)
                                p.Add($"<kbd>{t.Value}</kbd>");
                            else if (tf.Bold && tf.Italic)
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
                    case Drawing d:
                    {
                        var block = ProcessDrawing(renderer, d);
                        if (block != null)
                        {

                            // See if we're in a list.
                            if (document.Last() is MarkdownList theList)
                            {
                                var lastBlock = theList[^1];
                                lastBlock.Remove(blockOwner);
                                lastBlock.Add(block);
                            }
                            else
                            {
                                document.Remove(blockOwner);
                                document.Add(block);
                            }
                        }
                        break;
                    }
                }
            }
        }

        private static MarkdownBlock ProcessDrawing(IMarkdownRenderer renderer, Drawing d)
        {
            var p = d.Picture;
            if (p == null)
                return null;

            if (p.Extensions.Contains(VideoExtension.ExtensionId))
            {
                string videoUrl = ((VideoExtension) p.Extensions.Get(VideoExtension.ExtensionId)).Source;
                if (string.IsNullOrEmpty(videoUrl) || !videoUrl.ToLower().StartsWith("http"))
                {
                    // See if there's a hyperlink.
                    videoUrl = p.Hyperlink?.OriginalString;
                }

                return !string.IsNullOrEmpty(videoUrl) ? new BlockQuote($"[!VIDEO]({videoUrl})") : null;
            }

            // Get the filename for the image.
            var filename = p.Description;
            if (!string.IsNullOrEmpty(filename))
            {
                filename = Path.GetFileName(filename);
                if (string.IsNullOrEmpty(filename) 
                    || !Path.HasExtension(filename)
                    || Path.GetInvalidFileNameChars().Any(filename.Contains))
                    filename = null;
            }

            Image theImage;

            // If we have an SVG version, we'll use that.
            var svgExtension = (SvgExtension) p.Extensions.Get(SvgExtension.ExtensionId);
            if (svgExtension != null)
            {
                theImage = svgExtension.Image;
                if (string.IsNullOrEmpty(filename))
                    filename = theImage.FileName;
            }
            else
            {
                theImage = p.Image;
                if (string.IsNullOrEmpty(filename))
                    filename = p.FileName;
            }

            using var input = theImage.OpenStream();
            using var output = File.OpenWrite(Path.Combine(renderer.MediaFolder, filename));
            input.CopyTo(output);

            string imagePath = Path.Combine(renderer.RelativeMediaFolder, filename).Replace('\\','/');
            return new DocfxImage(d.Description, imagePath);
            //return new Julmar.GenMarkdown.Image(d.Description, imagePath);
        }
    }
}
