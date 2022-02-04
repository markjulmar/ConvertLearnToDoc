using System.Diagnostics;
using System.IO;
using System.Linq;
using DXPlus;
using GenMarkdown.DocFx.Extensions;
using Julmar.GenMarkdown;
using DXText = DXPlus.Text;
using Image = DXPlus.Image;
using Paragraph = DXPlus.Paragraph;
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
                    // Simple text
                    case DXText t:
                    {
                        var p = (Julmar.GenMarkdown.Paragraph) blockOwner;

                        if (element.Parent is Hyperlink hl)
                        {
                            p.Add(Text.Link(hl.Text, hl.Uri.ToString()));
                        }
                        else if (t.Value.Length > 0)
                        {
                            if (tf.KbdTag) AppendText(p, t.Value, "<kbd>", "</kbd>");
                            else if (tf.Bold && tf.Italic) AppendText<BoldItalicText>(p, t.Value);
                            else if (tf.Bold) AppendText<BoldText>(p, t.Value);
                            else if (tf.Italic) AppendText<ItalicText>(p, t.Value);
                            else if (tf.Monospace) AppendText<InlineCode>(p, t.Value);
                            else p.Add(t.Value);
                        }
                        break;
                    }
                    // Some kind of line break
                    case Break b:
                    {
                        //var p = (Julmar.GenMarkdown.Paragraph)blockOwner;
                        //if (b.Type == BreakType.Line)
                        //    p.Add(Text.LineBreak);
                        break;
                    }
                    // Picture/image/video
                    case Drawing d:
                    {
                        var block = ProcessDrawing(renderer, d, element);
                        if (block != null)
                        {
                            // See if we're in a list.
                            if (document.Last() is MarkdownList theList)
                            {
                                var lastBlock = theList[^1];
                                if (blockOwner.ToString().TrimEnd('\r','\n').Length == 0)
                                    lastBlock.Remove(blockOwner);
                                lastBlock.Add(block);
                            }
                            else
                            {
                                Debug.Assert(blockOwner.ToString().TrimEnd('\r', '\n').Length == 0);
                                document.Remove(blockOwner);
                                document.Add(block);
                            }
                        }
                        break;
                    }
                }
            }
        }

        private static void AppendText<T>(Julmar.GenMarkdown.Paragraph paragraph, string text) where T : Text
        {
            if (paragraph.LastOrDefault() is T li)
            {
                li.Text += text;
            }
            else
            {
                var t = typeof(T).Name switch
                {
                    nameof(BoldText) => new BoldText(text),
                    nameof(ItalicText) => new ItalicText(text),
                    nameof(BoldItalicText) => new BoldItalicText(text),
                    nameof(InlineCode) => new InlineCode(text),
                    _ => new Text(text)
                };
                paragraph.Add(t);
            }
        }

        private static void AppendText(Julmar.GenMarkdown.Paragraph paragraph, string text, string prefix, string suffix)
        {
            // Check to see if the _previous_ text has the same emphasis. If so, we'll add this.
            if (paragraph.LastOrDefault()?.Text.EndsWith(suffix) == true)
            {
                var lastItem = paragraph.Last();
                lastItem.Text += text;
            }
            else
            {
                paragraph.Add($"{prefix}{text}{suffix}");
            }
        }

        private static MarkdownBlock ProcessDrawing(IMarkdownRenderer renderer, Drawing d, Run runOwner)
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

                return !string.IsNullOrEmpty(videoUrl) ? new BlockQuote($"[!VIDEO {videoUrl}]") : null;
            }

            // Get the filename for the image. Pandoc stores it in the description, check there first.
            var filename = p.Description;
            if (!string.IsNullOrEmpty(filename))
            {
                filename = Path.GetFileName(filename);
                if (!IsPossibleFilename(filename))
                    filename = null;
            }

            // Next, see if there's a comment. We can strip off any prefix like "ImageFilename:"
            if (filename == null)
            {
                var comments = (runOwner.Parent as Paragraph)?.Comments?.ToList();
                if (comments?.Count > 0)
                {
                    foreach (var text in comments.SelectMany(c => c.Comment.Paragraphs.Select(pp => pp.Text))
                                                                 .Where(t => !string.IsNullOrEmpty(t))
                                                                 .Select(t => t.Trim()))
                    {
                        if (text.Length > 3 && !text.Contains('\n'))
                        {
                            for (int i = 0; i < text.Length; i++)
                            {
                                string fn = text.Substring(i, text.Length - i).Trim();
                                if (IsPossibleFilename(fn))
                                {
                                    filename = fn;
                                    break;
                                }
                            }
                        }
                    }
                }
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

            // Strip any parameters
            string suffix = string.Empty;
            int index = filename.IndexOf('#');
            if (index > 0)
            {
                suffix = filename[index..];
                filename = filename[..index];
            }

            using var input = theImage.OpenStream();
            using var output = File.OpenWrite(Path.Combine(renderer.MediaFolder, filename));
            input.CopyTo(output);

            bool border = p.BorderColor != null;
            string imagePath = Path.Combine(renderer.RelativeMediaFolder, filename).Replace('\\', '/') + suffix;

            var paragraph = d.Parent.Parent as Paragraph;
            bool isLightbox = paragraph?.Comments.Any(c => c.Comment.Paragraphs.Any(p => p.Text == "lightbox")) == true;

            if (border || d.IsDecorative || isLightbox)
            {
                return new DocfxImage(d.Description, imagePath) { Border = border, LocScope = d.IsDecorative ? "noloc" : null, Lightbox = imagePath+"#lightbox" };
            }

            return new Julmar.GenMarkdown.Image(d.Description, imagePath);
        }

        /// <summary>
        /// Test for a possible filename.
        /// </summary>
        /// <param name="text">Text</param>
        /// <returns>True/False</returns>
        private static bool IsPossibleFilename(string text) =>
            !string.IsNullOrEmpty(text)
            && Path.HasExtension(text)
            && !Path.GetInvalidPathChars().Any(text.Contains)
            && !Path.GetInvalidFileNameChars().Any(text.Contains);
    }
}
