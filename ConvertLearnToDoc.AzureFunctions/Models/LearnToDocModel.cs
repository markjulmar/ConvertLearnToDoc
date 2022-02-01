using System.Linq;

namespace ConvertLearnToDoc.AzureFunctions.Models
{
    public class LearnToDocModel
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

        public string Repository { get; set; }
        public string Branch { get; set; }
        public string Folder { get; set; }
        public string ZonePivot { get; set; }
        public bool EmbedNotebookData { get; set; }

        public bool IsValid()
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
                   && ValidRepos.Contains(Repository)
                   && !string.IsNullOrEmpty(Branch)
                   && !string.IsNullOrEmpty(Folder);
        }
    }
}
