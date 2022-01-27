using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace ConvertLearnToDocWeb.Models
{
    public class ConversionViewModel
    {
        // Learn to doc
        [Url(ErrorMessage = "Please specify a full URL to the Learn module page.")]
        public string ModuleUrl { get; set; }
        public string GithubRepo { get; set; }
        public string GithubBranch { get; set; }
        [RegularExpression("^(/[\\w-]+)+/*$", ErrorMessage = "Please specify the folder leading to the module index.yml.")]
        public string GithubFolder { get; set; }
        public string ZonePivot { get; set; }

        // Doc to learn
        public IFormFile WordDoc { get; set; }
        public bool UseAsterisksForBullets { get; set; }
        public bool UseAsterisksForEmphasis { get; set; }
        public bool OrderedListUsesSequence { get; set; }
        public bool UseAlternateHeaderSyntax { get; set; }
        public bool UseIndentsForCodeBlocks { get; set; }

    }
}
