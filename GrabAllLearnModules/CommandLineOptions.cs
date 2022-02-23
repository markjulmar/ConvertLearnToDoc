using CommandLine;

namespace CompareAll
{
    internal sealed class CommandLineOptions
    {
        [Value(0, Required = true)]
        public string InputFolder { get; set; }

        [Value(1, Required = false)]
        public string OutputFolder { get; set; }

        [Option('d', "Debug", HelpText = "Debug output, save temp files.")]
        public bool Debug { get; set; }

        [Option('t', "OutputType", HelpText = "Output type - Text, CSV, Markdown", Default = PrintType.Text)]
        public PrintType OutputType { get; set; }
    }
}
