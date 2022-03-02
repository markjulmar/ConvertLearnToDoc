using CommandLine;

namespace ConvertAll
{
    internal sealed class CommandLineOptions
    {
        [Value(0, Required = true)]
        public string InputFolder { get; set; }

        [Value(1, Required = true)]
        public string OutputFolder { get; set; }

        [Option('d', "Debug", HelpText = "Debug output, save temp files.")]
        public bool Debug { get; set; }
    }
}
