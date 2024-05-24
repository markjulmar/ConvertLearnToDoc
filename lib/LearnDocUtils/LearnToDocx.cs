using MSLearnRepos;

namespace LearnDocUtils;

public static class LearnToDocx
{
    public static async Task<List<string>> ConvertFromRepoAsync(string url, string organization, string repo, string branch, string folder,
        string outputFile, string accessToken = null, DocumentOptions options = null)
    {
        if (string.IsNullOrEmpty(organization)) 
            throw new ArgumentException($"'{nameof(organization)}' cannot be null or empty.", nameof(organization));
        if (string.IsNullOrEmpty(repo))
            throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
        if (string.IsNullOrEmpty(folder))
            throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

        return await Convert(LearnRepoService.Create(organization, repo, branch, accessToken), url, 
            folder, outputFile, options);
    }

    /*
    public static async Task<List<string>> ConvertFromFolderAsync(string learnFolder, string outputFile, DocumentOptions options = null)
    {
        if (string.IsNullOrWhiteSpace(learnFolder))
            throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

        if (!Directory.Exists(learnFolder))
            throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

        return await Convert(LearnRepoService.Create(learnFolder), learnFolder, outputFile, options);
    }
    */

    private static async Task<List<string>> Convert(ILearnRepoService learnRepo, string url,
        string moduleFolder, string docxFile, DocumentOptions options)
    {
        var rootTemp = options?.Debug == true ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
        var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
        while (Directory.Exists(tempFolder))
        {
            tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
        }

        // Download the module
        var (module, markdownFile) = await ModuleCombiner.DownloadModuleAsync(
            learnRepo, moduleFolder, tempFolder,
            options?.EmbedNotebookContent == true);

        try
        {
            // Convert the file.
            return await MarkdownToDocConverter.ConvertMarkdownToDocx(learnRepo, url, moduleFolder, module, markdownFile, docxFile, options?.ZonePivot, options?.Debug==true);
        }
        finally
        {
            if (options is {Debug: false})
            {
                Directory.Delete(tempFolder, true);
            }
        }
    }
}