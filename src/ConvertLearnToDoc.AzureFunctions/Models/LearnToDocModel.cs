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
        "learn-mec-pr",
    };

    public bool EmbedNotebookData { get; set; }

    public override bool IsValid()
    {
        string repository = Repository;
        if (repository.Contains('.'))
            repository = repository[..repository.IndexOf('.')];

        // Check mandatory input
        return base.IsValid() 
            && ValidRepos.Contains(repository);
    }
}