namespace ConvertLearnToDoc.AzureFunctions.Models;

public class SinglePageToDocModel
{
    public string Repository { get; set; }
    public string Branch { get; set; }
    public string Folder { get; set; }
    public string ZonePivot { get; set; }
    
    public virtual bool IsValid()
    {
        // Do some cleanup of input
        Repository = Repository?.Trim().ToLower();
        Branch = Branch?.Trim();
        Folder = Folder?.Trim();
        ZonePivot = ZonePivot?.Trim();
        if (ZonePivot == string.Empty) ZonePivot = null;
        if (string.IsNullOrEmpty(Branch)) Branch = "live";

        // Check mandatory input
        return !string.IsNullOrEmpty(Repository)
               && !string.IsNullOrEmpty(Branch)
               && !string.IsNullOrEmpty(Folder);
    }
}