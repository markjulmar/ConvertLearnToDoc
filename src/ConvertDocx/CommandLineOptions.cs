using CommandLine;

namespace ConvertDocx;

internal sealed class CommandLineOptions
{
    //[Option('i', "input", Required = true, HelpText = "Input file or folder.")]
    [Value(0, MetaName = "input file",
            Required = true, HelpText = "Input file or folder.")]
    public required string InputFile { get; set; }
        
    //[Option('o', "output", Required = true, HelpText = "Output file or folder.")]
    [Value(1, MetaName = "output file",
        Required = false, HelpText = "Output file or folder.")]
    public string? OutputFile { get; set; }

    [Option('s', "singlePage", HelpText = "Output should be a single page (Markdown file).")]
    public bool SinglePageOutput { get; set; }
        
    [Option('g', "Organization", HelpText = "GitHub organization")]
    public string? Organization { get; set; }

    [Option('r', "Repo", HelpText = "GitHub repo")]
    public string? GitHubRepo { get; set; }

    [Option('b', "Branch", HelpText = "GitHub branch, defaults to 'live'")]
    public string? GitHubBranch { get; set; }

    [Option('t', "Token", HelpText = "GitHub access token")]
    public string? AccessToken { get; set; }

    [Option('d', "Debug", HelpText = "Debug output, save temp files")]
    public bool Debug { get; set; }
        
    [Option('z', "ZonePivot", HelpText = "Zone pivot to render to doc, defaults to all")]
    public string? ZonePivot { get; set; }

    [Option('n', "Notebook", HelpText = "Convert notebooks into document, only used on MS Learn content")]
    public bool ConvertNotebooks { get; set; }

    [Option('p', "Plain", HelpText = "Prefer plain markdown (no docs extensions)")]
    public bool PreferPlainMarkdown { get; set; }

    [Option('x', "Ignore-Metadata", HelpText = "Ignore any existing metadata in the document. Only for training modules, use this to ensure a new module is created.")]
    public bool IgnoreMetadata { get; set;}

    [Option('u', "UseGenericIds", HelpText = "Generate filenames and UIDs based on generic pattern ('unit-xy')")]
    public bool UseGenericIds { get; set; }
    
    [Option('f', "OutputFormat",
        Default = OutputFormat.Docx,
        HelpText = "Output format for the conversion. Default is 'docx'.")]
    public OutputFormat OutputFormat { get;set; }
}