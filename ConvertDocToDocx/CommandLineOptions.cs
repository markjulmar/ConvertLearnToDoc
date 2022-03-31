using CommandLine;

namespace ConvertDocToDocx;

internal sealed class CommandLineOptions
{
    [Option('i', "input", Required = true, HelpText = "Input file or folder.")]
    public string InputFile { get; set; }
        
    [Option('o', "output", Required = true, HelpText = "Output file or folder.")]
    public string OutputFile { get; set; }
        
    [Option('g', "Organization", HelpText = "GitHub organization")]
    public string Organization { get; set; }

    [Option('r', "Repo", HelpText = "GitHub repo")]
    public string GitHubRepo { get; set; }

    [Option('b', "Branch", HelpText = "GitHub branch, defaults to 'live'.")]
    public string GitHubBranch { get; set; }

    [Option('t', "Token", HelpText = "GitHub access token.")]
    public string AccessToken { get; set; }

    [Option('d', "Debug", HelpText = "Debug output, save temp files.")]
    public bool Debug { get; set; }
        
    [Option('p', "Pivot", HelpText = "Zone pivot to render to doc")]
    public string ZonePivot { get; set; }
}