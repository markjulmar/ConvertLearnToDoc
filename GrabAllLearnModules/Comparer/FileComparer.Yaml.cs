using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using YamlDotNet.Serialization;

namespace CompareAll.Comparer;

public static partial class FileComparer
{
    public static IEnumerable<Difference> Yaml(string fn1, string fn2)
    {
        var yaml1 = ParseYaml(File.ReadAllText(fn1));
        var yaml2 = ParseYaml(File.ReadAllText(fn2));

        var result = CompareYamlData("", yaml1, yaml2);

        foreach (var difference in result)
            yield return difference;
    }

    private static IEnumerable<YamlDiff> CompareYamlData(string parentKey,
        IReadOnlyDictionary<object,object> yaml1, 
        IReadOnlyDictionary<object, object> yaml2)
    {
        foreach (var item in yaml1)
        {
            var original = item.Value;
            bool exists = yaml2.TryGetValue(item.Key, out var newValue);
            string key = string.IsNullOrEmpty(parentKey) ? item.Key.ToString() : $"{parentKey}:{item.Key}";

            if (!exists)
            {
                yield return new YamlDiff(key)
                {
                    Change = ChangeType.Deleted,
                    OriginalValue = ConvertToString(original)
                };
                
                continue;
            }

            if (original is List<object> l1
                && newValue is List<object> l2)
            {
                foreach (var child in CompareYamlList(key, l1, l2))
                    yield return child;
                continue;
            }

            if (original is Dictionary<object, object> o1
                && newValue is Dictionary<object, object> o2)
            {
                foreach (var child in CompareYamlData(key, o1, o2))
                    yield return child;
                
                continue;
            }

            var text1 = ConvertToString(original);
            var text2 = ConvertToString(newValue);

            if (text1 != text2)
            {
                yield return new YamlDiff(key)
                {
                    Change = ChangeType.Changed,
                    OriginalValue = text1,
                    NewValue = text2,
                };
            }
        }

        // Pick off unique keys in the second collection
        string[] keys1 = yaml1.Keys.Select(k => k.ToString()).ToArray();
        string[] keys2 = yaml2.Keys.Select(k => k.ToString()).ToArray();

        foreach (var key in keys2.Where(k => !keys1.Contains(k)))
        {
            yield return new YamlDiff($"{parentKey}{key}")
            {
                Change = ChangeType.Added,
                NewValue = ConvertToString(yaml2[key])
            };
        }
    }

    private static IEnumerable<YamlDiff> CompareYamlList(string parentKey, List<object> lstOriginal, List<object> lstNew)
    {
        int index = 0;
        for (; index < lstOriginal.Count; index++)
        {
            string key = parentKey + $".{index}";

            var original = lstOriginal[index];
            object newValue = lstNew.Count > index ? lstNew[index] : null;

            if (original != null && newValue == null)
            {
                yield return new YamlDiff(key)
                {
                    Change = ChangeType.Deleted,
                    OriginalValue = ConvertToString(original)
                };
                continue;
            }

            if (original is List<object> l1
                && newValue is List<object> l2)
            {
                foreach (var child in CompareYamlList(key, l1, l2))
                    yield return child;
                continue;
            }

            if (original is Dictionary<object, object> d1
                && newValue is Dictionary<object, object> d2)
            {
                foreach (var child in CompareYamlData(key, d1, d2))
                    yield return child;
                continue;
            }

            var text1 = ConvertToString(original);
            var text2 = ConvertToString(newValue);

            if (text1 != text2)
            {
                yield return new YamlDiff(key)
                {
                    Change = ChangeType.Changed,
                    OriginalValue = text1,
                    NewValue = text2,
                };
            }
        }

        for (; index < lstNew.Count; index++)
        {
            string key = parentKey + $".{index}";
            object value = lstNew[index];

            yield return new YamlDiff(key)
            {
                Change = ChangeType.Changed,
                NewValue = ConvertToString(value)
            };
        }
    }

    private static string ConvertToString(object value)
    {
        if (value is Dictionary<object, object> dict)
        {
            var sb = new StringBuilder("{ ");
            int count = 0;
            foreach (var item in dict.OrderBy(kvp => kvp.Key.ToString()))
            {
                if (count++ > 0)
                    sb.Append(", ");
                sb.Append(item.Key + ": ")
                  .Append(ConvertToString(item.Value));
            }

            sb.Append(" }");
            return sb.ToString();
        }

        if (value is List<object> lst)
        {
            var sb = new StringBuilder("[ ");
            sb.Append(string.Join(", ", lst.Select(ConvertToString)))
              .Append(" ]");
            return sb.ToString();
        }

        string text = value?.ToString() ?? "";
        if (long.TryParse(text, out var number)) return number.ToString();
        if (bool.TryParse(text, out var tf)) return tf.ToString();
        return $"\"{text}\"";
    }

    private static Dictionary<object, object> ParseYaml(string yaml)
    {
        // We lose the comments with the deserializer, so capture them
        // before running it through YamlDotNet.
        var comments = new List<string>();
        var sb = new StringBuilder();
        var reader = new StringReader(yaml);
        while (true)
        {
            var line = reader.ReadLine();
            if (line == null) break;
            if (line.FirstOrDefault() == '#')
            {
                comments.Add(line);
            }
            else
            {
                sb.AppendLine(line);
            }
        }

        var deserializer = new Deserializer();
        var dynamicYaml = deserializer.Deserialize<dynamic>(sb.ToString());

        var dictionary = (Dictionary<object, object>)dynamicYaml;
        for (int i = 0; i < comments.Count; i++)
            dictionary.Add($"comment{i}", comments[i]);

        return dictionary;
    }
}