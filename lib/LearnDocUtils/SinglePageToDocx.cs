using MSLearnRepos;

namespace LearnDocUtils
{
    public static class SinglePageToDocx
    {
        public static async Task<List<string>> ConvertFromRepoAsync(string url, string organization, string repo, string branch, 
            string inputFile, string outputFile, string accessToken = null, DocumentOptions options = null)
        {
            if (string.IsNullOrEmpty(organization))
                throw new ArgumentException($"'{nameof(organization)}' cannot be null or empty.", nameof(organization));
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(inputFile))
                throw new ArgumentException($"'{nameof(inputFile)}' cannot be null or empty.", nameof(inputFile));

            return await Convert(LearnRepoService.Create(organization, repo, branch, accessToken), url,
                inputFile, outputFile, options);
        }

        public static async Task<List<string>> ConvertFromFileAsync(string url, string inputFile, string outputFile, DocumentOptions options = null)
        {
            if (string.IsNullOrWhiteSpace(inputFile))
                throw new ArgumentException($"'{nameof(inputFile)}' cannot be null or whitespace.", nameof(inputFile));

            if (!File.Exists(inputFile))
                throw new DirectoryNotFoundException($"{inputFile} does not exist.");

            string folder = Path.GetDirectoryName(inputFile);
            if (string.IsNullOrEmpty(folder))
                folder = Directory.GetCurrentDirectory();

            return await Convert(LearnRepoService.Create(folder), url, inputFile, outputFile, options);
        }

        private static async Task<List<string>> Convert(ILearnRepoService learnRepo, string url,
            string inputFile, string docxFile, DocumentOptions options)
        {
            var rootTemp = options?.Debug == true ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            Directory.CreateDirectory(tempFolder);

            // Download the page
            var markdownFile = await DocsDownloader.DownloadFileAsync(learnRepo, inputFile, tempFolder);

            try
            {
                // Convert the file.
                return await MarkdownToDocConverter.ConvertMarkdownToDocx(learnRepo, url, Path.GetDirectoryName(inputFile),
                    null, markdownFile, docxFile, options?.ZonePivot, options?.Debug == true);
            }
            finally
            {
                if (options is { Debug: false })
                {
                    Directory.Delete(tempFolder, true);
                }
            }
        }
    }
}
