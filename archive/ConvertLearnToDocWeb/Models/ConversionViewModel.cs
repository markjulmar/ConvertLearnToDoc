using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace ConvertLearnToDocWeb.Models
{
    public class ConversionViewModel
    {
        public bool IsLearnToDoc { get; set; }

        // Learn to doc
        [Url(ErrorMessage = "Please specify a full URL to the Learn module page.")]
        public string ModuleUrl { get; set; }
        public string GitHubOrg { get; set; }
        public string GithubRepo { get; set; }
        public string GithubBranch { get; set; }
        public string GithubFolder { get; set; }
        public string ZonePivot { get; set; }
        public bool EmbedNotebookData { get; set; }
        public string TdRid { get; set; }

        // Doc to learn
        public IFormFile WordDoc { get; set; }
        public bool UseAsterisksForBullets { get; set; }
        public bool UsePlainMarkdown { get; set; }
        public bool UseAsterisksForEmphasis { get; set; }
        public bool OrderedListUsesSequence { get; set; }
        public bool UseAlternateHeaderSyntax { get; set; }
        public bool UseIndentsForCodeBlocks { get; set; }
        public bool PrettyPipeTables { get; set; }
        public string FdRid { get; set; }
        public bool UseGenericIds { get; set; }
        public bool IgnoreMetadata { get; set; }
    }
}
