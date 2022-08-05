using CommandLine;

namespace ConvertDocx;

internal sealed class CommandLineOptions
{
    [Option('i', "input", Required = true, HelpText = "Input file or folder.")]
    public string InputFile { get; set; }
        
    [Option('o', "output", Required = true, HelpText = "Output file or folder.")]
    public string OutputFile { get; set; }

    [Option('s', "singlePage", HelpText = "Output should be a single page (Markdown file).")]
    public bool SinglePageOutput { get; set; }
        
    [Option('g', "Organization", HelpText = "GitHub organization")]
    public string Organization { get; set; }

    [Option('r', "Repo", HelpText = "GitHub repo")]
    public string GitHubRepo { get; set; }

    [Option('b', "Branch", HelpText = "GitHub branch, defaults to 'live'")]
    public string GitHubBranch { get; set; }

    [Option('t', "Token", HelpText = "GitHub access token")]
    public string AccessToken { get; set; }

    [Option('d', "Debug", HelpText = "Debug output, save temp files")]
    public bool Debug { get; set; }
        
    [Option('z', "ZonePivot", HelpText = "Zone pivot to render to doc, defaults to all")]
    public string ZonePivot { get; set; }

    [Option('n', "Notebook", HelpText = "Convert notebooks into document, only used on MS Learn content")]
    public bool ConvertNotebooks { get; set; }

    [Option('p', "Plain", HelpText = "Prefer plain markdown (no docs extensions)")]
    public bool PreferPlainMarkdown { get; set; }
}