namespace ConvertLearnToDoc.AzureFunctions.Models;

public class PageToDocModel
{
    public string Organization { get; set; }
    public string Repository { get; set; }
    public string Branch { get; set; }
    public string Folder { get; set; }
    public string ZonePivot { get; set; }
    
    public virtual bool IsValid()
    {
        // Do some cleanup of input
        Organization = string.IsNullOrEmpty(Organization) ? MSLearnRepos.Constants.DocsOrganization : Organization.Trim();
        Repository = Repository?.Trim().ToLower();
        Branch = Branch?.Trim();
        Folder = Folder?.Trim();
        ZonePivot = ZonePivot?.Trim();
        if (ZonePivot == string.Empty) ZonePivot = null;
        if (string.IsNullOrEmpty(Branch)) Branch = "live";

        // Check mandatory input
        return !string.IsNullOrEmpty(Organization)
               && !string.IsNullOrEmpty(Repository)
               && !string.IsNullOrEmpty(Branch)
               && !string.IsNullOrEmpty(Folder);
    }
}