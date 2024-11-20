namespace ConvertLearnToDoc.Utility;

public class ContentRef
{
    public string Organization { get; set; }
    public string Repository { get; set; }
    public string Branch { get; set; }
    public string Folder { get; set; }
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