using GenMarkdown.DocFx.Extensions;

namespace Docx.Renderer.Markdown.Renderers;

public class RunRenderer : MarkdownObjectRenderer<Run>
{
    protected override void Render(IMarkdownRenderer renderer, 
        MarkdownDocument document, MarkdownBlock blockOwner, 
        Run element, RenderBag tags)
    {
        if (blockOwner == null)
        {
            blockOwner = new Paragraph();
            document.Add(blockOwner);
        }

        var tf = tags.Get<TextFormatting>(nameof(TextFormatting));
        if (element.Properties?.Bold == true)
            tf.Bold = true;
        if (element.Properties?.Italic == true)
            tf.Italic = true;
        if (element.Properties?.CapsStyle == CapsStyle.SmallCaps)
            tf.KbdTag = true;
        if (!string.IsNullOrEmpty(tf.StyleName))
            tf.StyleName = element.StyleName;
        if (TextFormatting.IsMonospaceFont(element.Properties?.Font))
            tf.Monospace = true;
        if (element.Properties?.Subscript == true)
            tf.Subscript = true;
        if (element.Properties?.Superscript == true)
            tf.Superscript = true;

        foreach (var e in element.Elements)
        {
            switch (e)
            {
                // Simple text
                case DXText t:
                {
                    var p = (Paragraph) blockOwner;

                    if (element.Parent is Hyperlink hl)
                    {
                        // Sometimes Word will split a hyperlink up across a Run boundary, we see it twice in a row due to the way
                        // I'm handling the links. This will catch that edge case and ignore the second appearance.
                        if (p.LastOrDefault() is InlineLink ll && ll.Url == hl.Uri.OriginalString && ll.Text == hl.Text)
                            continue;

                        p.Add(Text.Link(hl.Text, hl.Uri?.OriginalString ?? "#"));
                    }
                    else if (t.Value.Length > 0)
                    {
                        if (tf.KbdTag) AppendText(p, t.Value, "<kbd>", "</kbd>");
                        else if (tf.Bold && tf.Italic) AppendText<BoldItalicText>(p, t.Value);
                        else if (tf.Bold) AppendText<BoldText>(p, t.Value);
                        else if (tf.Italic) AppendText<ItalicText>(p, t.Value);
                        else if (tf.Monospace) AppendText<InlineCode>(p, t.Value);
                        else if (tf.Subscript) AppendText(p, t.Value, "<sub>", "</sub>");
                        else if (tf.Superscript) AppendText(p, t.Value, "<sup>", "</sup>");
                        else p.Add(ConvertSpecialCharacters(t.Value));
                    }
                    break;
                }
                // Some kind of line break
                case Break b:
                {
                    if (b.Type == BreakType.Line)
                    {
                        //var p = (Paragraph)blockOwner;
                        //p.Add(Text.LineBreak);
                    }
                    break;
                }
                // Picture/image/video
                case Drawing d:
                {
                    var block = ProcessDrawing(renderer, d, element);
                    if (block != null)
                    {
                        // See if we're in a list. If so, remove the empty paragraph from the 
                        // list, and then add the image to the list so it's indented.
                        if (document.Last() is MarkdownList theList)
                        {
                            var lastBlock = theList[^1];
                            if (blockOwner.ToString().TrimEnd('\r','\n').Length == 0)
                                lastBlock.Remove(blockOwner);
                            lastBlock.Add(block);
                        }
                        else if (document.Last() is Julmar.GenMarkdown.Table table)
                        {
                            var row = table.Last();
                            var cell = row.Last();
                            cell.Content = block;
                        }
                        // Remove the empty paragraph from the document and add the image instead.
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

    private static string ConvertSpecialCharacters(string text)
    {
        // TODO: if this gets larger, move to a dictionary.
        return text.Replace("❎", "[x]")
            .Replace("⬜", "[ ]");
    }

    private static void AppendText<T>(Paragraph paragraph, string text) where T : Text
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

    private static void AppendText(Paragraph paragraph, string text, string prefix, string suffix)
    {
        // Check to see if the _previous_ text has the same emphasis. If so, we'll add this.
        if (paragraph.LastOrDefault()?.Text.EndsWith(suffix) == true)
        {
            var lastItem = paragraph.Last();
            lastItem.Text += text;
        }
        else
        {
            paragraph.Add(new RawInline($"{prefix}{text}{suffix}"));
        }
    }

    private static MarkdownBlock ProcessDrawing(IMarkdownRenderer renderer, Drawing d, Run runOwner)
    {
        var p = d.Picture;
        if (p == null)
            return null;

        // Look for the video extension - if present, this is an embedded video.
        if (p.ImageExtensions.Contains(VideoExtension.ExtensionId))
        {
            string videoUrl = ((VideoExtension) p.ImageExtensions.Get(VideoExtension.ExtensionId)).Source;
            if (string.IsNullOrEmpty(videoUrl) || !videoUrl.ToLower().StartsWith("http"))
            {
                // See if there's a hyperlink.
                videoUrl = p.Hyperlink?.OriginalString;
            }
            return !string.IsNullOrEmpty(videoUrl) ? new BlockQuote($"[!VIDEO {videoUrl}]") : null;
        }

        bool extractMedia = true;

        // Get the filename for the image from the drawing properties, if null, use the picture name.
        string filename = d.Name ?? p.Name;
        string folder = null;

        if (!string.IsNullOrEmpty(filename))
        {
            // Word doesn't normally put folder names into the image name. Learn>Docx does though so we
            // can look for folder info here and see if this was an image referenced outside the module
            // folder - if so, we don't need to extract it.
            folder = Path.GetDirectoryName(filename);
            if (!string.IsNullOrEmpty(folder))
            {
                if (folder.StartsWith('/')
                    || folder != "." && !folder.StartsWith($"..{Path.DirectorySeparatorChar}media"))
                    extractMedia = false;
            }

            if (extractMedia)
            {
                filename = Path.GetFileName(filename);
                if (!IsPossibleFilename(filename))
                    filename = null;
            }
        }

        // Next, see if there's a comment. We can strip off any prefix like "ImageFilename:"
        if (filename == null)
        {
            var comments = (runOwner.Parent as DXParagraph)?.Comments?.ToList();
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

        DXImage theImage;

        // If we have an SVG version, we'll use that.
        var svgExtension = (SvgExtension)p.ImageExtensions.Get(SvgExtension.ExtensionId);
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

        // Extract the media if it wasn't originally in some other folder.
        if (extractMedia)
        {
            // Write file
            using var input = theImage.OpenStream();
            using var output = File.OpenWrite(Path.Combine(folder == "." ? renderer.MarkdownFolder : renderer.MediaFolder, filename));
            input.CopyTo(output);

            // Determine the URL for the Markdown content.
            folder = folder == "." ? "" : Path.GetRelativePath(renderer.MarkdownFolder, renderer.MediaFolder);
        }

        bool border = p.BorderColor != null;
        string imagePath = extractMedia
            ? Path.Combine(folder, filename).Replace('\\', '/') + suffix
            : filename;

        var paragraph = d.Parent.Parent as DXParagraph;

        // If the image has a caption, we'll use that -- otherwise see if we have a description.
        string description = null;
        if (paragraph?.NextParagraph?.Properties.StyleName == HeadingType.Caption.ToString())
        {
            description = paragraph.NextParagraph.Text;
            //TODO: loc? Remove figure prefix
            index = description.IndexOf(':');
            if (index > 0)
            {
                index++;
                while (description[index] == ' ') index++;
                description = description[index..];
            }
        }

        string altText = d.Description;
        bool useExtension = FindCommentValue(paragraph, "useExtension") != null;
        string locScope = FindCommentValue(paragraph, "loc-scope");
        string link = FindCommentValue(paragraph, "link");
        string lightboxUrl = FindCommentValue(paragraph, "lightbox");

        if (border || d.IsDecorative == true || lightboxUrl != null
            || !string.IsNullOrEmpty(locScope) || useExtension || !string.IsNullOrEmpty(link))
        {
            var image = new DocfxImage(altText, imagePath, description) { Border = border, Link = link };
            if (!string.IsNullOrEmpty(locScope))
                image.LocScope = locScope;
            else if (d.IsDecorative == true)
                image.LocScope = "other";
            if (lightboxUrl != null)
                image.Lightbox = lightboxUrl;
                
            return image;
        }

        return new Julmar.GenMarkdown.Image(altText, imagePath, description);
    }

    /// <summary>
    /// Look for a specific comment on the given paragraph. We assume the keys are space delimited
    /// and quoted if they have spaces. Surrounding quotes are removed. If the key has no value, then the
    /// key itself is returned. If the key doesn't exist, null is returned.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="prefix"></param>
    /// <returns></returns>
    private static string FindCommentValue(DXParagraph paragraph, string prefix)
    {
        if (paragraph != null)
        {
            var comments = paragraph.Comments.SelectMany(c => c.Comment.Paragraphs.Select(p => p.Text ?? ""));
            var found = comments.FirstOrDefault(c => c.Contains(prefix, StringComparison.InvariantCultureIgnoreCase));
            if (!string.IsNullOrEmpty(found))
            {
                found = found.TrimEnd('\r', '\n');

                int start = found.IndexOf(prefix, StringComparison.InvariantCultureIgnoreCase) + prefix.Length;
                if (start == found.Length || found[start] == ' ')
                    return prefix;

                Debug.Assert(found[start] == ':');
                start++;
                char lookFor = ' ';
                if (found[start] == '\"')
                {
                    lookFor = '\"';
                    start++;
                }

                int end = start;
                while (end < found.Length)
                {
                    if (found[end] == lookFor)
                        break;
                    end++;
                }

                Debug.Assert(start >= 0 && found.Length > start);
                return found.Substring(start, end - start);
            }
        }

        return null;
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