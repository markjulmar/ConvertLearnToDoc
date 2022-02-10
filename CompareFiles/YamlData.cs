namespace CompareFiles;

public class YamlData
{
    public int LineNumber { get; init; }
    public string Line { get; init; }
    public string Key { get; set; }
    public string Value { get; set; }
    public List<YamlData> Children { get; } = new();
    public override string ToString() => Key;
}