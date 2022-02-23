using System.Text;

namespace CompareAll.Comparer;

public class MarkdownDiff : Difference
{
    public int? OriginalPosition { get; init; }
    public int? NewPosition { get; init; }

    public override string Print(PrintType type)
    {
        var sb = new StringBuilder();

        if (type == PrintType.Text)
        {
            sb.AppendLine($"{Change} {OriginalPosition?.ToString() ?? "none"}/{NewPosition?.ToString() ?? "none"}: ");
            switch (Change)
            {
                case ChangeType.Deleted:
                    if (OriginalValue.Length > 0)
                        sb.AppendLine(ConvertNonPrintables(OriginalValue));
                    break;
                case ChangeType.Added:
                    if (NewValue.Length > 0)
                        sb.AppendLine(ConvertNonPrintables(NewValue));
                    break;
                case ChangeType.Changed:
                default:
                    sb.AppendLine($"O: {ConvertNonPrintables(OriginalValue)}")
                      .AppendLine($"N: {ConvertNonPrintables(NewValue)}");
                    break;
            }
        }
        else
        {
            sb.Append(Change).Append(',')
                .Append($"{OriginalPosition}/{NewPosition}").Append(',')
                .Append('\"').Append(ConvertNonPrintables(OriginalValue)).Append('\"').Append(',')
                .Append('\"').Append(ConvertNonPrintables(NewValue)).Append('\"').Append(',');
        }

        return sb.ToString();
    }
}