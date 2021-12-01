using System.ComponentModel.DataAnnotations;

namespace ConvertLearnToDocWeb.Models
{
    public class LearnToDocViewModel
    {
        [Url(ErrorMessage = "Please specify a full URL to the Learn module page.")]
        public string ModuleUrl { get; set; }
        public string GithubRepo { get; set; }
        public string GithubBranch { get; set; }
        [RegularExpression("^(/[\\w-]+)+/*$", ErrorMessage = "Please specify the folder leading to the module index.yml.")]
        public string GithubFolder { get; set; }
    }
}
