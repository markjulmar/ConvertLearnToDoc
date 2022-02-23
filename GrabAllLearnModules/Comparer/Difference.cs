using System.Text;

namespace CompareAll.Comparer;

public abstract class Difference
{
    public abstract string Key { get; }
    public ChangeType Change { get; set; }
    public string OriginalValue { get; init; }
    public string NewValue { get; init; }

    public string EscapedOriginalValue => ConvertNonPrintables(OriginalValue);
    public string EscapedNewValue => ConvertNonPrintables(NewValue);

    private static string ConvertNonPrintables(string text)
    {
        text ??= "";

        text = text.Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");

        return !text.StartsWith('\"') ? "\"" + text + "\"" : text;
    }

    public override string ToString()
    {
        var sb = new StringBuilder();

        sb.Append($"{Change} {Key}: ");
        switch (Change)
        {
            case ChangeType.Deleted:
                if (OriginalValue.Length > 0) sb.AppendLine(EscapedOriginalValue);
                break;
            case ChangeType.Added:
                if (NewValue.Length > 0) sb.AppendLine(EscapedNewValue);
                break;
            case ChangeType.Changed:
            default:
                sb.AppendLine()
                    .AppendLine($"  O: {EscapedOriginalValue}")
                    .AppendLine($"  N: {EscapedNewValue}");
                break;
        }

        return sb.ToString();
    }

}