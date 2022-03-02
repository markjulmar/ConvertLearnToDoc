using MSLearnRepos;

namespace LearnDocUtils;

public class ModuleMetadata
{
    private readonly Module moduleData;
    public Module ModuleData => moduleData;
    public ModuleMetadata(Module moduleData)
    {
        this.moduleData = moduleData ?? new Module();
        this.moduleData.Metadata ??= new Metadata();
    }

    public static string GetList(List<string> items, string defaultValue) =>
        items == null || items.Count == 0
            ? defaultValue
            : string.Join(Environment.NewLine, items.Select(s => "- " + s));
}