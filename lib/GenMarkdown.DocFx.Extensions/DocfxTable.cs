using System.Text;
using Julmar.GenMarkdown;

namespace GenMarkdown.DocFx.Extensions;

/// <summary>
/// This generates a Docfx table extension (:::row:::)
/// </summary>
public class DocfxTable : Table
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="columns"></param>
    public DocfxTable(int columns) : base(columns)
    {
    }

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="columns"></param>
    public DocfxTable(ColumnAlignment[] columns) : base(columns)
    {
    }

    /// <inheritdoc />
    public override void Write(TextWriter writer, MarkdownFormatting formatting)
    {
        var sb = new StringBuilder();
        foreach (var row in Children)
        {
            sb.AppendLine(":::row:::");

            for (int colIndex = 0; colIndex < row.Count; colIndex++)
            {
                var cell = row[colIndex];
                var sw = new StringWriter();
                cell.Content?.Write(sw, formatting);

                if (cell.ColumnSpan == 1)
                {
                    sb.AppendLine("    :::column:::");
                    sb.Append("    ");
                    sb.AppendLine(sw.ToString().TrimStart(' ').Trim('\r','\n'));
                    sb.AppendLine("    :::column-end:::");
                }
                else
                {
                    sb.AppendLine($"    :::column span=\"{cell.ColumnSpan}\":::");
                    sb.Append("    ");
                    sb.AppendLine(sw.ToString().TrimStart(' ').Trim('\r', '\n'));
                    sb.AppendLine("    :::column-end:::");
                    colIndex += cell.ColumnSpan - 1;
                }
            }

            sb.AppendLine(":::row-end:::");
        }

        writer.WriteLine(sb);
    }
}