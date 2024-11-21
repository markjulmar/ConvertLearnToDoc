using CommandLine;
using Julmar.DocsToMarkdown;
using LearnDocUtils;
using MSLearnRepos;
using System.Diagnostics;
using System.Reflection.Metadata;

namespace ConvertDocx;

public static class Program
{
    public static async Task Main(string[] args)
    {
        CommandLineOptions? options = null;
        new Parser(cfg => cfg.HelpWriter = Console.Error)
            .ParseArguments<CommandLineOptions>(args)
            .WithParsed(clo => options = clo);
        if (options == null) return;

        List<string> errors = [];

        try
        {
            if (options.InputFile!.StartsWith("http"))
            {
                if (RequiresGitHubToken(options))
                {
                    options.AccessToken = Environment.GetEnvironmentVariable("GITHUB_TOKEN")
                        ?? await GetGitHubToken();
                    if (string.IsNullOrEmpty(options.AccessToken))
                    {
                        await Console.Error.WriteLineAsync("GitHub access token is required for the specified options.");
                        return;
                    }
                }

                EnsureValidOutputFile(options);

                if (options.OutputFormat == OutputFormat.Docx)
                {
                    options.OutputFile = Path.ChangeExtension(options.OutputFile, ".docx");
                    errors = !string.IsNullOrEmpty(options.AccessToken) 
                        ? await ConvertFromRepoAsync(options) 
                        : await DownloadAndConvertAsync(options);
                }
                else
                {
                    options.OutputFile = Path.ChangeExtension(options.OutputFile, ".md");
                    await DownloadMarkdown(options);
                }

            }
            else if (options.InputFile.EndsWith(".docx", StringComparison.CurrentCultureIgnoreCase))
            {
                await ConvertDocxToMarkdown(options);
            }
            else
            {
                Console.WriteLine("ConvertDocx: unknown input file type.");
            }
        }
        catch (AggregateException aex)
        {
            var ex = aex.Flatten();
            errors.Add($"Conversion failed: {ex.Message} ({ex.GetType().Name})");
            errors.AddRange(ex.InnerExceptions.Select(e => e.Message));
        }
        catch (Exception ex)
        {
            errors.Add($"Conversion failed: {ex.Message} ({ex.GetType().Name})");
        }

        errors?.ForEach(Console.Error.WriteLine);
    }

    private static void EnsureValidOutputFile(CommandLineOptions options)
    {
        if (options.OutputFile == null)
        {
            var segments = options.InputFile.Split('/').Where(s => !string.IsNullOrEmpty(s)).ToArray();
            options.OutputFile = Path.ChangeExtension(segments[^1], options.OutputFormat == OutputFormat.Docx ? ".docx" : ".md");
        }

        if (Path.GetFileName(options.OutputFile).IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        {
            throw new ArgumentException($"Invalid characters in output filename: {options.OutputFile}.");
        }
    }

    private static bool RequiresGitHubToken(CommandLineOptions options)
    {
        return !string.IsNullOrEmpty(options.GitHubRepo) ||
               !string.IsNullOrEmpty(options.Organization) ||
               !string.IsNullOrEmpty(options.GitHubBranch);
    }

    private static async Task<List<string>> ConvertFromRepoAsync(CommandLineOptions options)
    {
        var metadata = await DocsMetadata.LoadFromUrlAsync(options.InputFile);
        if (metadata == null || metadata.ContentPath == null)
            throw new ArgumentException($"Unable to load metadata from {options.InputFile}.");

        options.Organization ??= metadata.Organization ?? Constants.DocsOrganization;
        options.GitHubRepo = metadata.Repository;
        options.GitHubBranch = metadata.Branch;

        var inputFile = metadata.PageType == "conceptual"
            ? metadata.ContentPath
            : Path.GetDirectoryName(metadata.ContentPath) 
              ?? throw new ArgumentException("Invalid metadata for URL.");

        if (inputFile.EndsWith(".md", StringComparison.InvariantCultureIgnoreCase))
        {
            Console.WriteLine($"Converting Docs article {options.InputFile} to {options.OutputFile}");
            return await SinglePageToDocx.ConvertFromRepoAsync(
                options.InputFile, options.Organization, options.GitHubRepo, options.GitHubBranch,
                inputFile, options.OutputFile, options.AccessToken, new DocumentOptions
                {
                    Debug = options.Debug,
                    ZonePivot = options.ZonePivot
                });
        }
        else
        {
            Console.WriteLine($"Converting Learn module {options.InputFile} to {options.OutputFile}");
            return await LearnToDocx.ConvertFromRepoAsync(
                options.InputFile, options.Organization, options.GitHubRepo, options.GitHubBranch,
                inputFile, options.OutputFile, options.AccessToken, new DocumentOptions
                {
                    Debug = options.Debug,
                    ZonePivot = options.ZonePivot,
                    EmbedNotebookContent = options.ConvertNotebooks
                });
        }
    }

    private static async Task DownloadMarkdown(CommandLineOptions options)
    {
        Console.WriteLine($"Converting {options.InputFile} to {options.OutputFile}");

        var tempFolder = Path.Combine(Path.GetTempPath(), "LearnDocs");
        if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
        Directory.CreateDirectory(tempFolder);

        try
        { 
            var downloader = new DocsConverter(tempFolder, new Uri(options.InputFile));
            var createdFiles = await downloader.ConvertAsync(!options.PreferPlainMarkdown, 
#if DEBUG
                tag => Console.Error.WriteLine($"Skipped: {tag.TrimStart().Substring(0, 20)}"));
#else
                null);
#endif

            if (createdFiles.Count == 0)
                throw new InvalidOperationException("No files created during download.");

            // Move the files to the output folder.
            var isModule = createdFiles.Any(f => f.Filename.EndsWith(".yml"));
            var inputFile = isModule
                ? createdFiles.First(f => f.FileType == FileType.Folder).Filename
                : createdFiles.Single(f => f.FileType == FileType.Markdown).Filename;


            if (isModule)
            {
                var outputFolder = Path.Combine(Path.ChangeExtension(options.OutputFile, "") ?? Directory.GetCurrentDirectory(), Path.GetFileName(inputFile));
                MoveOrCopyDirectory(inputFile, outputFolder);
                Console.WriteLine($"Module \"{options.InputFile}\" downloaded to {outputFolder}");
            }
            else
            {
                var outputFolder = Path.GetDirectoryName(options.OutputFile) ?? Directory.GetCurrentDirectory();
                File.Move(inputFile, options.OutputFile!);
    
                if (createdFiles.Count > 1)
                {
                    foreach (var file in createdFiles.Where(f => f.FileType == FileType.Image || f.FileType == FileType.Asset))
                    {
                        var relativeFile = Path.GetRelativePath(tempFolder, file.Filename);
                        var destFile = Path.Combine(outputFolder, relativeFile);
                        File.Move(file.Filename, destFile);
                    }
                }

                Console.WriteLine($"Article \"{options.InputFile}\" downloaded to {options.OutputFile}");
            }

        }
        finally
        {
            if (Directory.Exists(tempFolder)) 
                Directory.Delete(tempFolder, true);
        }

    }

    private static void MoveOrCopyDirectory(string inputFolder, string outputFolder)
    {
        inputFolder = Path.GetFullPath(inputFolder);
        outputFolder = Path.GetFullPath(outputFolder);

        try
        {
            Directory.Move(inputFolder, outputFolder);
            return;
        }
        catch (IOException)
        {
        }
        
        CopyFolder(inputFolder, outputFolder);
    }

    private static void CopyFolder(string inputFolder, string outputFolder)
    {
        // Copy the files.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Copy all files
        foreach (string file in Directory.GetFiles(inputFolder))
        {
            string destFile = Path.Combine(outputFolder, Path.GetFileName(file));
            File.Copy(file, destFile, true);
        }

        // Recursively copy all subdirectories
        foreach (string dir in Directory.GetDirectories(inputFolder))
        {
            string destDir = Path.Combine(outputFolder, Path.GetFileName(dir));
            CopyFolder(dir, destDir);
        }
    }

    private static async Task<List<string>> DownloadAndConvertAsync(CommandLineOptions options)
    {
        Console.WriteLine($"Downloading {options.InputFile}");

        var tempFolder = Path.Combine(Path.GetTempPath(), "LearnDocs");
        if (Directory.Exists(tempFolder)) Directory.Delete(tempFolder, true);
        Directory.CreateDirectory(tempFolder);

        try
        { 
            var downloader = new DocsConverter(tempFolder, new Uri(options.InputFile));
            var createdFiles = await downloader.ConvertAsync(!options.PreferPlainMarkdown, 
#if DEBUG
                tag => Console.Error.WriteLine($"Skipped: {tag.TrimStart().Substring(0, 20)}"));
#else
                null);
#endif

            if (createdFiles.Count == 0)
                throw new InvalidOperationException("No files created during download.");

            var isModule = createdFiles.Any(f => f.Filename.EndsWith(".yml"));
            var inputFile = isModule
                ? createdFiles.First(f => f.FileType == FileType.Folder).Filename
                : createdFiles.Single(f => f.FileType == FileType.Markdown).Filename;

            if (!isModule)
            {
                Console.WriteLine($"Converting Docs article {inputFile} to {options.OutputFile}");
                return await SinglePageToDocx.ConvertFromFileAsync(options.InputFile, inputFile, options.OutputFile, 
                    new DocumentOptions {
                        Debug = options.Debug,
                        ZonePivot = options.ZonePivot
                    });
            }
            else
            {
                Console.WriteLine($"Converting Learn module {inputFile} to {options.OutputFile}");
                return await LearnToDocx.ConvertFromFolderAsync(options.InputFile, inputFile, options.OutputFile, 
                    new DocumentOptions {
                        Debug = options.Debug,
                        ZonePivot = options.ZonePivot,
                        EmbedNotebookContent = options.ConvertNotebooks
                    });
            }
        }
        finally
        {
            if (Directory.Exists(tempFolder)) 
                Directory.Delete(tempFolder, true);
        }
    }

    private static async Task ConvertDocxToMarkdown(CommandLineOptions options)
    {
        if (options.SinglePageOutput)
        {
            options.OutputFile ??= Path.ChangeExtension(options.InputFile, ".md");
            Console.WriteLine($"Converting {options.InputFile} to single-page Markdown {options.OutputFile}");
            await DocxToSinglePage.ConvertAsync(options.InputFile, options.OutputFile, 
                new MarkdownOptions { Debug = options.Debug, UsePlainMarkdown = options.PreferPlainMarkdown });
        }
        else
        {
            options.OutputFile ??= Path.ChangeExtension(options.InputFile, "");
            Console.WriteLine($"Converting {options.InputFile} to Learn module {options.OutputFile}");
            await DocxToLearn.ConvertAsync(options.InputFile, options.OutputFile,
                new MarkdownOptions { Debug = options.Debug, UsePlainMarkdown = options.PreferPlainMarkdown });
        }
    }

    private static async Task<string?> GetGitHubToken()
    {
        var psi = new ProcessStartInfo("op")
        {
            Arguments = "read \"op://personal/github.com/token\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        try
        {
            using var process = Process.Start(psi);
            if (process != null)
            {
                var token = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync();
                return token.TrimEnd();
            }
        }
        catch (Exception ex)
        {
#if DEBUG
            Console.Error.WriteLine($"Failed to retrieve GitHub token: {ex.Message}");
#endif
        }

        return null;
    }
}
