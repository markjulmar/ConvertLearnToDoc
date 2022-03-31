using System.Text;
using System.Xml;
using DXPlus;

namespace LearnDocUtils;

/// <summary>
/// Class used to dump the contents of a Word document.
/// </summary>
internal class DocxDebug
{
    private readonly StringBuilder builder = new();

    /// <summary>
    /// Dumps the known object structure of the document
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    public static string Dump(IDocument document)
    {
        var dumper = new DocxDebug();
        foreach (var block in document.Blocks)
            dumper.DumpBlock(block, 0);
        return dumper.builder.ToString();
    }

    /// <summary>
    /// Dumps the raw XML for the document
    /// </summary>
    /// <param name="docx"></param>
    /// <returns></returns>
    public static string FormatXml(IDocument docx)
    {
        var document = new XmlDocument();
        document.Load(new StringReader(docx.RawDocument()));

        var builder = new StringBuilder();
        using (var writer = new XmlTextWriter(new StringWriter(builder)))
        {
            writer.Formatting = System.Xml.Formatting.Indented;
            document.Save(writer);
        }

        return builder.ToString();
    }

    private void DumpBlock(Block block, int level)
    {
        if (block is Paragraph p)
            DumpParagraph(p, level);
        else if (block is Table t)
            DumpTable(t, level);
        else if (block is UnknownBlock ub)
        {
            string prefix = new(' ', level * 3);
            builder.AppendLine($"{prefix}{ub.Name}");
        }
    }

    private void DumpParagraph(Paragraph block, int level)
    {
        string prefix = new(' ', level * 3);

        string listInfo = "";
        if (block.IsListItem())
        {
            listInfo = $"{block.GetNumberingFormat()} {block.GetListLevel()} #{block.GetListIndex() + 1} ";
        }

        builder.AppendLine($"{prefix}p: {block.Id} StyleName=\"{block.Properties.StyleName}\" {listInfo}{DumpObject(block.Properties.DefaultFormatting)}");
        foreach (var run in block.Runs)
        {
            DumpRun(run, level + 1);
        }

        foreach (var comment in block.Comments)
        {
            DumpCommentRef(comment, level + 1);
        }
    }

    private void DumpCommentRef(CommentRange comment, int level)
    {
        string prefix = new(' ', level * 3);

        builder.AppendLine($"{prefix}Comment id={comment.Comment.Id} by {comment.Comment.AuthorName} ({comment.Comment.AuthorInitials})");
        builder.AppendLine($"{prefix}   > start: {comment.RangeStart?.Text}, end: {comment.RangeEnd?.Text}");
            
        foreach (var p in comment.Comment.Paragraphs)
        {
            DumpParagraph(p, level + 1);
        }
    }

    private void DumpTable(Table table, int level)
    {
        string prefix = new(' ', level * 3);
        builder.AppendLine($"{prefix}tbl Design={table.Design} {table.CustomTableDesignName} {table.ConditionalFormatting}");
        foreach (var row in table.Rows)
        {
            DumpRow(row, level + 1);
        }
    }

    private void DumpRow(TableRow row, int level)
    {
        string prefix = new(' ', level * 3);
        builder.AppendLine($"{prefix}tr");
        foreach (var cell in row.Cells)
        {
            DumpCell(cell, level + 1);
        }
    }

    private void DumpCell(TableCell cell, int level)
    {
        string prefix = new(' ', level * 3);
        builder.AppendLine($"{prefix}tc");
        foreach (var p in cell.Paragraphs)
        {
            DumpParagraph(p, level + 1);
        }
    }

    private void DumpRun(Run run, int level)
    {
        string prefix = new(' ', level * 3);

        var parent = run.Parent;
        if (parent is Hyperlink hl)
        {
            builder.AppendLine($"{prefix}hyperlink: {hl.Id} <{hl.Uri}> \"{hl.Text}\"");
            prefix += "   ";
            level++;
        }

        builder.AppendLine($"{prefix}r: {DumpObject(run.Properties)}");
        foreach (var item in run.Elements)
        {
            DumpRunElement(item, level + 1);
        }
    }

    private void DumpRunElement(ITextElement item, int level)
    {
        string prefix = new(' ', level * 3);

        string text = "";
        switch (item)
        {
            case Text t:
                text = "\"" + t.Value + "\"";
                break;
            case Break b:
                text = b.Type + "Break";
                break;
            case CommentRef cr:
                text = $"{cr.Id} - {string.Join(". ", cr.Comment.Paragraphs.Select(p => p.Text))}";
                break;
            case Drawing d:
            {
                text = $"{prefix}{item.ElementType}: Id={d.Id} ({Math.Round(d.Width, 0)}x{Math.Round(d.Height, 0)}) - {d.Name}: \"{d.Description}\"";
                if (d.Hyperlink != null)
                {
                    text += $", Hyperlink=\"{d.Hyperlink.OriginalString}\"";
                }

                var p = d.Picture;
                if (p != null)
                {
                    text += $"{Environment.NewLine}{prefix}   pic: Id={p.Id}, Rid=\"{p.RelationshipId}\" {p.FileName} ({Math.Round(p.Width??0, 0)}x{Math.Round(p.Height??0, 0)}) - {p.Name}: \"{p.Description}\"";
                    if (p.Hyperlink != null)
                    {
                        text += $", Hyperlink=\"{p.Hyperlink.OriginalString}\"";
                    }

                    if (p.BorderColor != null)
                    {
                        text += $", BorderColor={p.BorderColor}";
                    }

                    string captionText = d.GetCaption();
                    if (captionText != null)
                    {
                        text += $", Caption=\"{captionText}\"";
                    }

                    foreach (var ext in p.ImageExtensions)
                    {
                        if (ext is SvgExtension svg)
                        {
                            text += $"{Environment.NewLine}{prefix}      SvgId={svg.RelationshipId} ({svg.Image.FileName})";
                        }
                        else if (ext is VideoExtension video)
                        {
                            text += $"{Environment.NewLine}{prefix}      Video=\"{video.Source}\" H={video.Height}, W={video.Width}";
                        }
                        else if (ext is DecorativeImageExtension dix)
                        {
                            text += $"{Environment.NewLine}{prefix}      DecorativeImage={dix.Value}";
                        }
                        else if (ext is LocalDpiExtension dpi)
                        {
                            text += $"{Environment.NewLine}{prefix}      LocalDpiOverride={dpi.Value}";
                        }
                        else
                        {
                            text += $"{Environment.NewLine}{prefix}      Extension {ext.UriId}";
                        }
                    }
                }

                builder.AppendLine(text);
                text = null;
                break;
            }
        }

        if (text != null)
            builder.AppendLine($"{prefix}{item.ElementType}: {text}");
    }

    private static string DumpObject(object obj)
    {
        var sb = new StringBuilder();
        Type t = obj.GetType();

        sb.Append($"{t.Name}: [");
        for (var index = 0; index < t.GetProperties().Length; index++)
        {
            var pi = t.GetProperties()[index];
            object val = pi.GetValue(obj);
            if (val != null)
            {
                if (index > 0) sb.Append(", ");
                sb.Append($"{pi.Name}={val}");
            }
        }

        sb.Append(']');
        return sb.ToString();
    }
}