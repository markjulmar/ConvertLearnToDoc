using CommandLine;
using CompareAll.Comparer;

namespace CompareAll
{
    internal sealed class CommandLineOptions
    {
        [Value(0, Required = true)]
        public string InputFolder { get; set; }
        
        [Option('d', "Debug", HelpText = "Debug output, save temp files.")]
        public bool Debug { get; set; }

        [Option('t', "OutputType", HelpText = "Output type - Text, CSV", Default = PrintType.Csv)]
        public PrintType OutputType { get; set; }
    }
}
