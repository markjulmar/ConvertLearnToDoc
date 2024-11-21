using System.Text;
using GenMarkdown.DocFx.Extensions;
using Image = Julmar.GenMarkdown.Image;

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

        // Pickup any formatting from the paragraph style - only do this for custom styles.
        if (!string.IsNullOrEmpty(tf.StyleName))
        {
            var style = element.Owner.Styles.Find(tf.StyleName, StyleType.Paragraph);
            if (style is {IsCustom: true})
            {
                if (style is { Formatting.Bold: true } || TextFormatting.IsBoldFont(style.Formatting?.Font))
                    tf.Bold = true;
                if (style is { Formatting.Italic: true })
                    tf.Italic = true;
                if (style is { Formatting.CapsStyle: CapsStyle.SmallCaps })
                    tf.KbdTag = true;
                if (TextFormatting.IsMonospaceFont(style.Formatting?.Font))
                    tf.Monospace = true;
                if (style is { Formatting.Subscript: true })
                    tf.Subscript = true;
                if (style is { Formatting.Superscript: true })
                    tf.Superscript = true;
            }
        }

        // Pickup any formatting from the run style
        var runStyle = element.StyleName;
        if (!string.IsNullOrEmpty(runStyle))
        {
            var style = element.Owner.Styles.Find(runStyle, StyleType.Character);
            if (style != null)
            {
                if (style is {Formatting.Bold: true} || TextFormatting.IsBoldFont(style.Formatting?.Font))
                    tf.Bold = true;
                if (style is {Formatting.Italic: true})
                    tf.Italic = true;
                if (style is {Formatting.CapsStyle: CapsStyle.SmallCaps})
                    tf.KbdTag = true;
                if (TextFormatting.IsMonospaceFont(style.Formatting?.Font))
                    tf.Monospace = true;
                if (style is {Formatting.Subscript:true})
                    tf.Subscript = true;
                if (style is {Formatting.Superscript:true})
                    tf.Superscript = true;
            }
        }

        // Pickup any formatting from the run itself
        if (element.Properties?.Bold == true || TextFormatting.IsBoldFont(element.Properties?.Font))
            tf.Bold = true;
        if (element.Properties?.Italic == true)
            tf.Italic = true;
        if (element.Properties?.CapsStyle == CapsStyle.SmallCaps)
            tf.KbdTag = true;
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
                        var url = renderer.ConvertAbsoluteUrl(hl.Uri?.OriginalString ?? "#");

                        // Sometimes Word will split a hyperlink up across a Run boundary, we see it twice in a row due to the way
                        // I'm handling the links. This will catch that edge case and ignore the second appearance.
                        if (p.LastOrDefault() is InlineLink ll && (hl.Uri == null || ll.Url == hl.Uri.OriginalString || ll.Url == url) && ll.Text == hl.Text)
                            continue;

                        // Remove any bold/italic spaces before this.
                        RenderHelpers.CollapseEmptyTags(p);

                        var link = new InlineLink(hl.Text, url);
                        if (tf.Bold) link.Bold = true;
                        if (tf.Italic) link.Italic = true;

                        p.Add(link);
                    }
                    else if (t.Value.Length > 0)
                    {
                        if (tf.KbdTag) AppendText(p, t.Value, "<kbd>", "</kbd>");
                        else if (tf is { Bold: true, Italic: true }) AppendText<BoldItalicText>(p, t.Value);
                        else if (tf.Bold) AppendText<BoldText>(p, t.Value);
                        else if (tf.Italic) AppendText<ItalicText>(p, t.Value);
                        else if (tf.Monospace) AppendText<InlineCode>(p, t.Value);
                        else if (tf.Subscript) AppendText(p, t.Value, "<sub>", "</sub>");
                        else if (tf.Superscript) AppendText(p, t.Value, "<sup>", "</sup>");
                        else
                        {
                            RenderHelpers.CollapseEmptyTags(p);
                            p.Add(new Text(ConvertSpecialCharacters(t.Value)));
                        }
                    }
                    break;
                }
                // Some kind of line break
                case Break b:
                {
                    if (b.Type == BreakType.Line)
                    {
                        var p = (Paragraph)blockOwner;
                        p.Add(new RawInline("<br>"));
                    }
                    break;
                }
                // Picture/image/video
                case Drawing d:
                {
                    var block = ProcessDrawing(renderer, d, element);
                    if (block != null)
                    {
                        if (d.Hyperlink != null && block is Image image)
                        {
                            block = new ImageLink(d.Hyperlink.OriginalString, image.AltText, image.ImagePath,
                                image.Description);
                        }

                        // See if we're in a list. If so, remove the empty paragraph from the 
                        // list, and then add the image to the list so it's indented.
                        if (document.LastOrDefault() is MarkdownList theList)
                        {
                            var lastBlock = theList[^1];
                            if (blockOwner.ToString().TrimEnd('\r','\n').Length == 0)
                                lastBlock.Remove(blockOwner);

                            // If this block doesn't have anything in it, then we should remove it
                            // and move up to the last non-empty block.
                            while (lastBlock.Count == 0 && theList.Count > 1)
                            {
                                theList.Remove(lastBlock);
                                lastBlock = theList[^1];
                            }
                            
                            lastBlock.Add(block);
                        }
                        else if (document.LastOrDefault() is Julmar.GenMarkdown.Table table)
                        {
                            var row = table.LastOrDefault();
                            if (row == null)
                                table.Add(new TableRow());

                            var cell = row?.LastOrDefault();
                            if (cell == null)
                            {
                                row.Add(block);
                            }
                            else
                            {
                                cell.Content = block;
                            }
                        }
                        // Remove the empty paragraph from the document and add the image instead.
                        else
                        {
                            if (blockOwner.ToString().TrimEnd('\r', '\n').Length == 0)
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
        var replacements = new Dictionary<char, string>
        {
            { '❎', "[x]" },
            { '⬜', "[ ]" },
            { '‛', "'" },
            { '’', "'" },
            { '“', "\"" },
            { '”', "\"" },
            { '–', "-" },
            { (char)0xa0, " " }
        };

        var sb = new StringBuilder(text.Length);
        for (var index = 0; index < text.Length; index++)
        {
            var c = text[index];
            sb.Append(replacements.TryGetValue(c, out var replacement)
                ? replacement
                : c);
        }

        return sb.ToString();
    }

    private static void AppendText<T>(Paragraph paragraph, string text) where T : Text
    {
        text = ConvertSpecialCharacters(text);

        if (paragraph.LastOrDefault() is T li)
        {
            li.Text += text;
        }
        else
        {
            if (paragraph.Count > 0 && 
                paragraph.Last()?.GetType() != typeof(T))
            {
                // Tag mismatch -- collapse any spaces before this.
                RenderHelpers.CollapseEmptyTags(paragraph);
            }

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
        text = ConvertSpecialCharacters(text);
        
        // Check to see if the _previous_ text has the same emphasis. If so, we'll add this.
        if (paragraph.LastOrDefault()?.Text.EndsWith(suffix) == true)
        {
            var lastItem = paragraph.Last();
            lastItem.Text += text;
        }
        else
        {
            // It's different -- collapse any tags.
            RenderHelpers.CollapseEmptyTags(paragraph);
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
            var videoUrl = ((VideoExtension) p.ImageExtensions.Get(VideoExtension.ExtensionId))?.Source;
            if (string.IsNullOrEmpty(videoUrl) || !videoUrl.ToLower().StartsWith("http"))
            {
                // See if there's a hyperlink.
                videoUrl = p.Hyperlink?.OriginalString;
            }

            if (string.IsNullOrEmpty(videoUrl))
                return null;

            if (renderer.PreferPlainMarkdown)
            {
                return new RawBlock(
                   $"<video width=\"320\" height=\"240\" controls>{Environment.NewLine}"
                      +$"   <source src=\"{videoUrl}\" type=\"video/mp4\">{Environment.NewLine}"
                      +$"   Your browser does not support the video tag.{Environment.NewLine}"
                      + "</video>"
                );
            }

            return new BlockQuote($"[!VIDEO {videoUrl}]");
        }

        bool extractMedia = true;

        // Get the filename for the image from the drawing properties, if null, use the picture name.
        string filename = d.Name ?? p.Name;
        string folder = string.Empty;

        if (!string.IsNullOrEmpty(filename))
        {
            // Word doesn't normally put folder names into the image name. Learn>Docx does though so we
            // can look for folder info here and see if this was an image referenced outside the local
            // folder - if so, we don't need to extract it.
            folder = Path.GetDirectoryName(filename);
            if (!string.IsNullOrEmpty(folder))
            {
                if (folder.StartsWith('/')
                    || folder.Split('/').Count(s => s == "..") > 1)
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
            var comments = (runOwner.Parent as DXParagraph)?.Comments.ToList();
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
            if (string.IsNullOrEmpty(filename) && theImage != null)
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
        int index = filename?.IndexOf('#') ?? -1;
        if (index > 0)
        {
            suffix = filename![index..];
            filename = filename[..index];
        }

        // Extract the media if it wasn't originally in some other folder.
        if (extractMedia && theImage != null)
        {
            string createPath;

            if (folder == ".")
            {
                createPath = renderer.MarkdownFolder;
            }
            else if (string.IsNullOrEmpty(folder) || folder!.Contains("media"))
            {
                createPath = renderer.MediaFolder;
            }
            else
            {
                createPath = Path.Combine(renderer.MarkdownFolder, folder!);
            }

            Debug.Assert(!string.IsNullOrEmpty(createPath));
            if (!Directory.Exists(createPath))
                Directory.CreateDirectory(createPath);

            // Write file
            using var input = theImage.OpenStream();
            using var output = File.OpenWrite(Path.Combine(createPath, filename));
            input.CopyTo(output);

            // Determine the URL for the Markdown content.
            folder = folder == "." ? string.Empty : Path.GetRelativePath(renderer.MarkdownFolder, createPath);
        }

        bool border = p.BorderColor != null;
        string imagePath = extractMedia
            ? Path.Combine(folder!, filename!).Replace('\\', '/') + suffix
            : filename;

        var paragraph = d.Parent?.Parent as DXParagraph;

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
        bool useExtension = FindCommentValue(paragraph, Globals.UseExtension) != null;
        string locScope = FindCommentValue(paragraph, "loc-scope");
        string link = FindCommentValue(paragraph, "link");
        string lightboxUrl = FindCommentValue(paragraph, "lightbox");

        if (!renderer.PreferPlainMarkdown && (border || d.IsDecorative == true || lightboxUrl != null
            || !string.IsNullOrEmpty(locScope) || useExtension || !string.IsNullOrEmpty(link)))
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

        return new Image(altText, imagePath, description);
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