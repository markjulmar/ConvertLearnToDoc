using MSLearnRepos;

namespace LearnDocUtils;

public class ModuleMetadata
{
    public Module ModuleData { get; }

    public ModuleMetadata(Module moduleData)
    {
        this.ModuleData = moduleData ?? new Module();
        this.ModuleData.Metadata ??= new Metadata();
    }

    public static string GetList(string header, List<string> items)
    {
        if (items == null || items.Count == 0) 
            return null;

        return $"{header}:" + Environment.NewLine 
                            + string.Join(Environment.NewLine, items.Select(s => "- " + s));
    }

    public static string GetOrCreateList(List<string> items, string defaultValue) =>
        items == null || items.Count == 0
            ? defaultValue
            : string.Join(Environment.NewLine, items.Select(s => "- " + s));
}