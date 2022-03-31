using System.ComponentModel.Design;
using CommandLine;
using LearnDocUtils;

namespace ConvertDocToDocx;

public static class Program
{
    static async Task Main(string[] args)
    {
        CommandLineOptions options = null;
        new Parser(cfg => { cfg.HelpWriter = Console.Error; })
            .ParseArguments<CommandLineOptions>(args)
            .WithParsed(clo => options = clo);
        if (options == null)
            return; // bad arguments or help.

        Console.WriteLine($"Doc/Docx: converting {Path.GetFileName(options.InputFile)}");

        List<string> errors = null;

        try
        {
            // Input is a Docs URL
            if (options.InputFile!.StartsWith("http"))
            {
                errors = await SinglePageToDocx.ConvertFromUrlAsync(options.InputFile,
                    options.OutputFile, options.AccessToken,
                    new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });
            }

            // Input is a repo + branch + file
            else if (!string.IsNullOrEmpty(options.GitHubRepo))
            {
                if (string.IsNullOrEmpty(options.Organization))
                    options.Organization = MSLearnRepos.Constants.DocsOrganization;

                errors = await SinglePageToDocx.ConvertFromRepoAsync(options.Organization, options.GitHubRepo, options.GitHubBranch,
                    options.InputFile, options.OutputFile, options.AccessToken, 
                    new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });

            }

            // Input is a local file for a docs page
            else if (File.Exists(options.InputFile))
            {
                string fileExtension = Path.GetExtension(options.InputFile)?.ToLower();

                if (fileExtension == ".md")
                {
                    errors = await SinglePageToDocx.ConvertFromFileAsync(options.InputFile, options.OutputFile,
                        new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });
                }
                // Input is a docx file
                else if (fileExtension == ".docx")
                {
                    if (string.IsNullOrEmpty(options.OutputFile))
                        options.OutputFile = Path.ChangeExtension(options.InputFile, "");
                    await DocxToSinglePage.ConvertAsync(options.InputFile, options.OutputFile, new MarkdownOptions { Debug = options.Debug });
                }
                
                else throw new InvalidOperationException($"Unknown input file type: {options.InputFile}.");
            }
        }
        catch (AggregateException aex)
        {
            throw aex.Flatten();
        }

        errors?.ForEach(Console.Error.WriteLine);

    }
}

