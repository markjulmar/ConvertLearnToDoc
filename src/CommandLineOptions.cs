using CommandLine;

namespace ConvertLearnToDoc
{
    internal sealed class CommandLineOptions
    {
        [Option('i', "input", Required = true)]
        public string InputFileOrFolder { get; set; }
        
        [Option('o', "output")]
        public string OutputFileOrFolder { get; set; }
        
        [Option('z', "zipOutput")]
        public bool ZipOutput { get; set; }

        [Option('r', "Repo")]
        public string GitHubRepo { get; set; }

        [Option('b', "Branch")]
        public string GitHubBranch { get; set; }

        [Option('t', "Token")]
        public string AccessToken { get; set; }
    }
}
