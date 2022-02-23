namespace CompareAll.Comparer;

public class MarkdownDiff : Difference
{
    public int? OriginalPosition { get; init; }
    public int? NewPosition { get; init; }

    public override string Key =>
        OriginalPosition != null && NewPosition != null ? $"{OriginalPosition.Value}/{NewPosition.Value}" :
        OriginalPosition != null ? OriginalPosition.Value.ToString() :
        NewPosition != null ? NewPosition.Value.ToString() : "";
}