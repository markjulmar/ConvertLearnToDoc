using CommandLine;
using Julmar.DocsToMarkdown;
using LearnDocUtils;
#if USE_GITHUB_PAT
using MSLearnRepos;
using System.Diagnostics;
#endif

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
        var tempFolder = Path.Combine(Path.GetTempPath(), "LearnDocs");

#if USE_GITHUB_PAT
        // Try to get a token.
        options.AccessToken ??= await GetGitHubToken;
        if (options.AccessToken == "") options.AccessToken = null;
#endif

        try
        {
            // Input is a Docs or Learn URL -> DOCX
            if (options.InputFile!.StartsWith("http"))
            {
                if (Path.GetFileName(options.OutputFile).IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                {
                    Console.Error.WriteLine($"Invalid characters in output filename: {options.OutputFile}.");
                    return;
                }

#if USE_GITHUB_PAT
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
#else
                // Download the article or training module.
                Console.WriteLine($"Downloading {options.InputFile}");
                if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
                Directory.CreateDirectory(tempFolder);

                var downloader = new DocsConverter(tempFolder, new Uri(options.InputFile));
                var createdFiles = await downloader.ConvertAsync(logger: tag => Console.Error.WriteLine($"Skipped: {tag.TrimStart().Substring(0,20)}"));
                if (createdFiles.Count == 0)
                {
                    Console.Error.WriteLine("No files created.");
                    return;
                }

                bool isModule = createdFiles.Any(f => f.Filename.EndsWith(".yml"));
                var url = options.InputFile;
                options.InputFile = isModule
                    ? createdFiles.First(f => f.FileType == FileType.Folder).Filename
                    : createdFiles.Single(f => f.FileType == FileType.Markdown).Filename;

                if (!Path.HasExtension(options.OutputFile))
                    options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");

                if (options.InputFile!.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine(
                        $"ConvertDocx: converting Docs Markdown {options.InputFile} to {options.OutputFile}");
                    errors = await SinglePageToDocx.ConvertFromFileAsync(url, 
                        options.InputFile, options.OutputFile,
                        new DocumentOptions { Debug = options.Debug, ZonePivot = options.ZonePivot });
                }
                else
                {
                    Console.WriteLine(
                        $"ConvertDocx: converting Learn module {options.InputFile} to {options.OutputFile}");
                    errors = await LearnToDocx.ConvertFromFolderAsync(url, 
                        options.InputFile, options.OutputFile,
                        new DocumentOptions
                        {
                            Debug = options.Debug, ZonePivot = options.ZonePivot,
                            EmbedNotebookContent = options.ConvertNotebooks
                        });
                }
#endif
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
        finally
        {
            if (Directory.Exists(tempFolder)) 
                Directory.Delete(tempFolder, true);
        }

        errors?.ForEach(Console.Error.WriteLine);
    }

#if USE_GITHUB_PAT
    static async Task<string> GetGitHubToken()
    {
        var psi = new ProcessStartInfo("op")
        {
            Arguments = "read \"op://personal/github.com/token\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(psi);
        if (process == null)
            return null;

        var token = await process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync();

        return token?.TrimEnd();
    }
#endif
}

