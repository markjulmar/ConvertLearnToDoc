namespace CompareAll.Comparer;

public abstract class Difference
{
    public ChangeType Change { get; set; }
    public string OriginalValue { get; set; }
    public string NewValue { get; set; }

    public override string ToString() => Print(PrintType.Text);

    protected static string ConvertNonPrintables(string text)
    {
        text ??= "";
        return text.Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }

    public abstract string Print(PrintType type);

}