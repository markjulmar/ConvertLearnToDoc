using System.Text;
using MSLearnRepos;

namespace LearnDocUtils;

/// <summary>
/// Representation of the unit metadata
/// </summary>
public class UnitMetadata
{
    public string Title { get; }
    public ModuleUnit Metadata { get; }

    public List<string> Lines { get; } = new();
    public bool HasContent => string.IsNullOrEmpty(Metadata.Notebook) && Lines.Count > 0 && Lines.Any(s => !string.IsNullOrWhiteSpace(s));

    public UnitMetadata(string title, ModuleUnit metadata)
    {
        Title = title;
        Metadata = metadata ?? new ModuleUnit();
        Metadata.Metadata ??= new MSLearnRepos.UnitMetadata();
    }

    public string BuildInteractivityOptions()
    {
        var sb = new StringBuilder();
        if (Metadata.UsesSandbox)
            sb.AppendLine("sandbox: true");
        if (!string.IsNullOrEmpty(Metadata.InteractivityType))
            sb.AppendLine($"interactive: {Metadata.InteractivityType}");
        if (Metadata.LabId > 0)
            sb.AppendLine($"labId: {Metadata.LabId}");
        if (!string.IsNullOrEmpty(Metadata.Notebook))
            sb.AppendLine($"notebook: {Metadata.Notebook}");
            
        return sb.ToString();
    }
}