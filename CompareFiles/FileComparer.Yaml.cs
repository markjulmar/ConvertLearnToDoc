namespace CompareFiles;

public static partial class FileComparer
{
    public static IEnumerable<Difference<YamlData>> Yaml(string fn1, string fn2)
    {
        string originalYaml1 = File.ReadAllText(fn1);
        string originalYaml2 = File.ReadAllText(fn2);

        originalYaml1 = originalYaml1.Replace("azureSandbox", "sandbox", StringComparison.InvariantCultureIgnoreCase);
        originalYaml2 = originalYaml2.Replace("azureSandbox", "sandbox", StringComparison.InvariantCultureIgnoreCase);

        var yaml1 = ParseYaml(originalYaml1);
        var yaml2 = ParseYaml(originalYaml2);

        foreach (var difference in CompareYamlData(fn1, fn2, yaml1, yaml2)) 
            yield return difference;
    }

    private static IEnumerable<Difference<YamlData>> CompareYamlData(string fn1, string fn2, IReadOnlyList<YamlData> yaml1, IReadOnlyList<YamlData> yaml2)
    {
        var emptyList = new List<YamlData>();

        foreach (var item in yaml1)
        {
            var item2 = yaml2.SingleOrDefault(k => k.Key == item.Key);
            if (item2 == null || item2.Value != item.Value)
            {
                yield return new Difference<YamlData> {
                    Filename1 = fn1, Filename2 = fn2,
                    Value1 = item, Value2 = item2,
                    LineNumber = item?.LineNumber
                };
            }

            if (item!.Children.Count > 0)
            {
                foreach (var child in CompareYamlData(fn1, fn2, item.Children,item2?.Children ?? emptyList))
                    yield return child;
            }
        }

        // Pick off unique keys in the second collection
        foreach (var item in yaml2.Where(kvp => yaml1.SingleOrDefault(it => it.Key == kvp.Key) == null))
        {
            yield return new Difference<YamlData>
            {
                Filename1 = fn1, Filename2 = fn2,
                Value1 = null, Value2 = item,
                LineNumber = item?.LineNumber
            };
        }
    }

    private static List<YamlData> ParseYaml(string yaml)
    {
        List<YamlData> data = new();

        using var sr = new StringReader(yaml);

        string line = sr.ReadLine();
        int lineNumber = 1;
        while (line != null)
        {
            // Skip whitespace
            if (string.IsNullOrWhiteSpace(line)) 
                continue;

            bool isRoot = line[0] != ' ';
            int index = line.IndexOf(':');

            if (line.StartsWith("#"))
            {
                var currentItem = new YamlData { LineNumber = lineNumber, Line = line };
                data.Add(currentItem);
                currentItem.Key = line;
            }
            else if (index > 0 && isRoot && line.TrimStart().Take(index).All(ch => ch != ' '))
            {
                var currentItem = new YamlData {LineNumber = lineNumber};
                data.Add(currentItem);

                currentItem.Key = line[..index].Trim();
                currentItem.Value = line[(index + 1)..].Trim();
            }
            else
            {
                var child = new YamlData { LineNumber = lineNumber, Line = line };

                // Don't split on colon if parent has a continuation block.
                if (data.Last().Value.Contains("|"))
                    index = -1;

                if (index > 0)
                {
                    child.Key = line[..index].Trim();
                    child.Value = line[(index + 1)..].Trim();
                }
                else
                {
                    child.Key = line.Trim();
                }
                data.Last().Children.Add(child);
            }

            line = sr.ReadLine();
            lineNumber++;
        }

        return data;
    }
}