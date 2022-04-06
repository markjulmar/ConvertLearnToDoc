using CommandLine;
using LearnDocUtils;
using MSLearnRepos;

namespace ConvertDocx;

public static class Program
{
    public static async Task Main(string[] args)
    {
        CommandLineOptions options = null;
        new Parser(cfg => { cfg.HelpWriter = Console.Error; })
            .ParseArguments<CommandLineOptions>(args)
            .WithParsed(clo => options = clo);
        if (options == null)
            return; // bad arguments or help.

        Console.WriteLine($"ConvertDocx: converting {options.InputFile} to {options.OutputFile}");

        List<string> errors = null;

        try
        {
            // Input is a Docs or Learn URL -> DOCX
            if (options.InputFile!.StartsWith("http"))
            {
                var metadata = await DocsMetadata.LoadFromUrlAsync(options.InputFile);
                
                options.Organization = metadata.Organization;
                options.GitHubRepo = metadata.Repository;
                options.GitHubBranch = metadata.Branch;
                options.InputFile = metadata.PageType == "conceptual" ? metadata.ContentPath : Path.GetDirectoryName(metadata.ContentPath);
            }

            // Input is a repo + branch + file -> DOCX
            if (!string.IsNullOrEmpty(options.GitHubRepo))
            {
                if (string.IsNullOrEmpty(options.Organization))
                    options.Organization = Constants.DocsOrganization;

                if (options.InputFile!.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                {
                    errors = await SinglePageToDocx.ConvertFromRepoAsync(
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });
                }
                else
                {
                    if (Path.HasExtension(options.InputFile))
                    {
                        options.InputFile = Path.GetDirectoryName(options.InputFile);
                    }

                    errors = await LearnToDocx.ConvertFromRepoAsync(
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot, EmbedNotebookContent = options.ConvertNotebooks });
                }
            }

            // Input is a local file for a docs page or .docx file
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
                    {
                        options.OutputFile = options.SinglePageOutput 
                            ? Path.ChangeExtension(options.InputFile, ".md") 
                            : Path.GetFileNameWithoutExtension(options.InputFile);
                    }

                    if (options.OutputFile!.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                        options.SinglePageOutput = true;

                    if (options.SinglePageOutput)
                    {
                        await DocxToSinglePage.ConvertAsync(options.InputFile, options.OutputFile, 
                            new MarkdownOptions { Debug = options.Debug });
                    }
                    else
                    {
                        await DocxToLearn.ConvertAsync(options.InputFile, options.OutputFile,
                            new MarkdownOptions { Debug = options.Debug });
                    }

                    if (options.ZipOutput && Directory.Exists(options.OutputFile))
                    {
                        string baseFolder = Path.GetDirectoryName(options.OutputFile);
                        if (string.IsNullOrEmpty(baseFolder))
                            baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                        string zipFile = Path.Combine(baseFolder,
                            Path.ChangeExtension(Path.GetFileNameWithoutExtension(options.OutputFile), "zip"));

                        System.IO.Compression.ZipFile.CreateFromDirectory(options.OutputFile, zipFile);
                    }
                }

                else throw new InvalidOperationException($"Unknown input file type: {options.InputFile}.");
            }

            // Input is a local folder (Learn module)
            else if (Directory.Exists(options.InputFile))
            {
                await LearnToDocx.ConvertFromFolderAsync(options.InputFile, options.OutputFile, 
                    new DocumentOptions { 
                        Debug = options.Debug, 
                        EmbedNotebookContent = options.ConvertNotebooks, 
                        ZonePivot = options.ZonePivot
                    });
            }
        }
        catch (AggregateException aex)
        {
            throw aex.Flatten();
        }

        errors?.ForEach(Console.Error.WriteLine);

    }
}

