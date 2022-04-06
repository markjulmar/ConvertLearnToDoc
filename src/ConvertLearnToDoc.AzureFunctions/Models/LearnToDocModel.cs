using System.Linq;

namespace ConvertLearnToDoc.AzureFunctions.Models;

public class LearnToDocModel : PageToDocModel
{
    private static readonly string[] ValidRepos =
    {
        "learn-sandbox-pr",
        "learn-sandbox",
        "learn-pr",
        "learn-dynamics-pr",
        "learn-bizapps-pr",
        "learn-m365-pr",
    };

    public bool EmbedNotebookData { get; set; }

    public override bool IsValid()
    {
        // Check mandatory input
        return base.IsValid() && ValidRepos.Contains(Repository);
    }
}