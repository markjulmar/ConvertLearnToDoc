using CommandLine;
using ConvertLearnToDoc;
using LearnDocUtils;

CommandLineOptions options = null;
new Parser(cfg => { cfg.HelpWriter = Console.Error; })
    .ParseArguments<CommandLineOptions>(args)
    .WithParsed(clo => options = clo);
if (options == null)
    return -1; // bad arguments or help.

Console.WriteLine($"Learn/Docx: converting {Path.GetFileName(options.InputFileOrFolder)}");

try
{
    // Input is a Learn module URL
    if (options.InputFileOrFolder!.StartsWith("http"))
    {
        var errors = await LearnToDocx.ConvertFromUrlAsync(options.InputFileOrFolder,
            options.OutputFileOrFolder, options.AccessToken, 
            new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks, ZonePivot = options.ZonePivot });
        errors.ForEach(Console.Error.WriteLine);
    }

    // Input is a repo + folder + branch
    else if (!string.IsNullOrEmpty(options.GitHubRepo))
    {
        if (string.IsNullOrEmpty(options.Organization))
            options.Organization = MSLearnRepos.Constants.DocsOrganization;

        await LearnToDocx.ConvertFromRepoAsync(options.Organization, options.GitHubRepo, options.GitHubBranch,
            options.InputFileOrFolder, options.OutputFileOrFolder, options.AccessToken, new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks, ZonePivot = options.ZonePivot });

    }

    // Input is a local folder containing a Learn module
    else if (Directory.Exists(options.InputFileOrFolder))
    {
        await LearnToDocx.ConvertFromFolderAsync(options.InputFileOrFolder, options.OutputFileOrFolder, new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks, ZonePivot = options.ZonePivot });
    }
    
    // Input is a docx file
    else
    {
        if (string.IsNullOrEmpty(options.OutputFileOrFolder))
            options.OutputFileOrFolder = Path.ChangeExtension(options.InputFileOrFolder, "");

        await DocxToLearn.ConvertAsync(options.InputFileOrFolder, options.OutputFileOrFolder, new MarkdownOptions { Debug = options.Debug });

        if (options.ZipOutput)
        {
            string baseFolder = Path.GetDirectoryName(options.OutputFileOrFolder);
            if (string.IsNullOrEmpty(baseFolder))
                baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string zipFile = Path.Combine(baseFolder,
                Path.ChangeExtension(Path.GetFileNameWithoutExtension(options.OutputFileOrFolder), "zip"));

            System.IO.Compression.ZipFile.CreateFromDirectory(options.OutputFileOrFolder, zipFile);
        }
    }
}
catch (AggregateException aex)
{
    throw aex.Flatten();
}

return 0;
