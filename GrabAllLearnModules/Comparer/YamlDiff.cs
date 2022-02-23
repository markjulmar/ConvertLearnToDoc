namespace CompareAll.Comparer;

public class YamlDiff : Difference
{
    private readonly string key;
    public YamlDiff(string key) { this.key = key; }
    public override string Key => key;
}