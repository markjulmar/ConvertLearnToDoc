using CommandLine;

namespace ConvertLearnToDoc
{
    internal sealed class CommandLineOptions
    {
        [Option('i', "input", Required = true)]
        public string InputFileOrFolder { get; set; }
        [Option('o', "output", Required = false)]
        public string OutputFileOrFolder { get; set; }
        [Option('z', "zipOutput", Required = false)]
        public bool ZipOutput { get; set; }
    }
}
