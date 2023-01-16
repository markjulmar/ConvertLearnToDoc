using System.ComponentModel.DataAnnotations;

namespace ConvertLearnToDoc.Shared;

public class ContentRef
{
    [MaxLength(200)]
    public string Organization { get; set; }
    
    [Required, MaxLength(200)]
    public string Repository { get; set; }
    
    [MaxLength(255)]
    public string Branch { get; set; }
    
    [Required, MaxLength(4096)]
    public string Folder { get; set; }
    
    public string? ZonePivot { get; set; }
    public bool EmbedNotebooks { get; set; }
    public string PageType { get; set; }

    public ContentRef()
    {
        Organization = Repository = Branch = Folder = PageType = string.Empty;
    }

    public override string ToString()
    {
        return $"{Organization}/{Repository}/{Branch}/{Folder}";
    }

    public bool IsValid()
    {
        // Do some cleanup of input
        Organization = string.IsNullOrEmpty(Organization) ? "MicrosoftDocs" : Organization.Trim();
        Repository = Repository.Trim().ToLower();
        Branch = Branch.Trim();
        Folder = Folder.Trim();
        ZonePivot = ZonePivot?.Trim();
        if (ZonePivot == string.Empty) ZonePivot = null;
        if (string.IsNullOrEmpty(Branch)) Branch = "live";

        // Must be Learn module or conceptual article.
        if (string.Compare(PageType, "conceptual", StringComparison.InvariantCultureIgnoreCase) != 0
            && string.Compare(PageType, "learn.module", StringComparison.InvariantCultureIgnoreCase) != 0)
            return false;

        // Check mandatory input
        return !string.IsNullOrEmpty(Organization)
               && !string.IsNullOrEmpty(Repository)
               && !string.IsNullOrEmpty(Branch)
               && !string.IsNullOrEmpty(Folder);
    }
}