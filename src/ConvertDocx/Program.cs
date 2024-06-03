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

        List<string> errors = null;

        try
        {
            // Input is a Docs or Learn URL -> DOCX
            if (options.InputFile!.StartsWith("http"))
            {
                var metadata = await DocsMetadata.LoadFromUrlAsync(options.InputFile);

                options.Organization = metadata.Organization ?? Constants.DocsOrganization;
                options.GitHubRepo = metadata.Repository;
                options.GitHubBranch = metadata.Branch;
                options.InputFile = metadata.PageType == "conceptual"
                    ? metadata.ContentPath
                    : Path.GetDirectoryName(metadata.ContentPath);

                if (!Path.HasExtension(options.OutputFile))
                    options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");

                if (options.InputFile!.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine(
                        $"ConvertDocx: converting Docs Markdown {options.InputFile} to {options.OutputFile}");
                    errors = await SinglePageToDocx.ConvertFromRepoAsync(options.InputFile,
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });
                }
                else
                {
                    if (Path.HasExtension(options.InputFile))
                        options.InputFile = Path.GetDirectoryName(options.InputFile);

                    Console.WriteLine(
                        $"ConvertDocx: converting Learn module {options.InputFile} to {options.OutputFile}");
                    errors = await LearnToDocx.ConvertFromRepoAsync(options.InputFile,
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions
                        {
                            Debug = options.Debug, ZonePivot = options.ZonePivot,
                            EmbedNotebookContent = options.ConvertNotebooks
                        });
                }
            }

            // Input is a Word document to Markdown
            else if (options.InputFile!.EndsWith(".docx", StringComparison.CurrentCultureIgnoreCase))
            {
                if (options.SinglePageOutput)
                {
                    Console.WriteLine(
                        $"DocxToSinglePage: converting {options.InputFile} to {options.OutputFile}");
                    await DocxToSinglePage.ConvertAsync(options.InputFile, options.OutputFile,
                        new MarkdownOptions { Debug = options.Debug, UsePlainMarkdown = options.PreferPlainMarkdown });
                }
                else
                {
                    Console.WriteLine(
                        $"DocxToLearn: converting {options.InputFile} to {options.OutputFile}");
                    await DocxToLearn.ConvertAsync(options.InputFile, options.OutputFile,
                        new MarkdownOptions { Debug = options.Debug, UsePlainMarkdown = options.PreferPlainMarkdown });
                }

            }
            else
            {
                Console.WriteLine("ConvertDocx: unknown input file type.");
            }
        }
        catch (AggregateException aex)
        {
            aex = aex.Flatten();
            await Console.Error.WriteLineAsync($"Conversion failed. Error: {aex.Message}");
#if DEBUG
            await Console.Error.WriteLineAsync(aex.ToString());
#endif
        }
        catch (Exception ex)
        {
            await Console.Error.WriteLineAsync($"Conversion failed. Error: {ex.Message}");
#if DEBUG
            await Console.Error.WriteLineAsync(ex.ToString());
#endif
        }

        errors?.ForEach(Console.Error.WriteLine);
    }
}

