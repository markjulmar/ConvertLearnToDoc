using Microsoft.AspNetCore.Http;

namespace ConvertLearnToDoc.AzureFunctions.Models
{
    public class DocToLearnModel
    {
        public IFormFile WordDoc { get; set; }
        public bool UseAsterisksForBullets { get; set; }
        public bool UseAsterisksForEmphasis { get; set; }
        public bool OrderedListUsesSequence { get; set; }
        public bool UseAlternateHeaderSyntax { get; set; }
        public bool UseIndentsForCodeBlocks { get; set; }
    }
}
