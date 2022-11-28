using CommandLine;
using Docx.Renderer.Markdown;
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

                options.Organization = metadata.Organization;
                options.GitHubRepo = metadata.Repository;
                options.GitHubBranch = metadata.Branch;
                options.InputFile = metadata.PageType == "conceptual"
                    ? metadata.ContentPath
                    : Path.GetDirectoryName(metadata.ContentPath);
            }

            // Input is a repo + branch + file -> DOCX
            if (!string.IsNullOrEmpty(options.GitHubRepo))
            {
                if (string.IsNullOrEmpty(options.Organization))
                    options.Organization = Constants.DocsOrganization;

                if (!Path.HasExtension(options.OutputFile))
                    options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");

                if (options.InputFile!.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine($"ConvertDocx: converting Docs Markdown {options.InputFile} to {options.OutputFile}");
                    errors = await SinglePageToDocx.ConvertFromRepoAsync(
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions {Debug = options.Debug, ZonePivot = options.ZonePivot});
                }
                else
                {
                    if (Path.HasExtension(options.InputFile))
                        options.InputFile = Path.GetDirectoryName(options.InputFile);

                    Console.WriteLine($"ConvertDocx: converting Learn module {options.InputFile} to {options.OutputFile}");
                    errors = await LearnToDocx.ConvertFromRepoAsync(
                        options.Organization, options.GitHubRepo, options.GitHubBranch,
                        options.InputFile, options.OutputFile, options.AccessToken,
                        new DocumentOptions
                        {
                            Debug = options.Debug, ZonePivot = options.ZonePivot,
                            EmbedNotebookContent = options.ConvertNotebooks
                        });
                }
            }

            // Input is a local file for a docs page or .docx file
            else if (File.Exists(options.InputFile))
            {
                string fileExtension = Path.GetExtension(options.InputFile)?.ToLower();
                if (fileExtension == ".md")
                {
                    if (!Path.HasExtension(options.OutputFile))
                        options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");

                    Console.WriteLine($"ConvertDocx: converting Docs Markdown {options.InputFile} to {options.OutputFile}");
                    errors = await SinglePageToDocx.ConvertFromFileAsync(options.InputFile, options.OutputFile,
                        new DocumentOptions {Debug = options.Debug, ZonePivot = options.ZonePivot});
                }
                // Input is a docx file
                else if (fileExtension == ".docx")
                {
                    if (options.OutputFile.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                        options.SinglePageOutput = true;
                    else if (options.SinglePageOutput && !Path.HasExtension(options.OutputFile))
                        options.OutputFile = Path.ChangeExtension(options.OutputFile, ".md");

                    Console.WriteLine($"ConvertDocx: converting Word document {options.InputFile} to {options.OutputFile}");
                    if (options.SinglePageOutput)
                    {
                        await DocxToSinglePage.ConvertAsync(options.InputFile, options.OutputFile,
                            new MarkdownOptions {Debug = options.Debug}, options.PreferPlainMarkdown);
                    }
                    else
                    {
                        await DocxToLearn.ConvertAsync(options.InputFile, options.OutputFile,
                            new LearnMarkdownOptions
                            {
                                Debug = options.Debug, 
                                IgnoreMetadata = options.IgnoreMetadata, 
                                UseGenericIds = options.UseGenericIds
                            });
                    }
                }

                else
                {
                    await Console.Error.WriteLineAsync(
                        $"Unknown input file type: {options.InputFile}. Must be URL, Markdown file, Word (.docx), or Learn folder.");
                    return;
                }
            }

            // Input is a local folder (Learn module)
            else if (Directory.Exists(options.InputFile))
            {
                Console.WriteLine($"ConvertDocx: converting Learn module {options.InputFile} to {options.OutputFile}");
                await LearnToDocx.ConvertFromFolderAsync(options.InputFile, options.OutputFile,
                    new DocumentOptions
                    {
                        Debug = options.Debug,
                        EmbedNotebookContent = options.ConvertNotebooks,
                        ZonePivot = options.ZonePivot
                    });
            }
            else
            {
                await Console.Error.WriteLineAsync(
                    $"Unknown input file/folder: {options.InputFile} does not exist.");
                return;
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

