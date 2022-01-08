using System;
using MSLearnRepos;
using System.IO;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class LearnToDocx
    {
        public static async Task ConvertFromUrlAsync(
            string url, string outputFile, string zonePivot, string accessToken,
            bool debug, IMarkdownToDocx converter)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(url);
            await ConvertFromRepoAsync(repo, branch, folder, outputFile, zonePivot, accessToken, debug, converter);
        }

        public static async Task ConvertFromRepoAsync(string repo, string branch, string folder,
                string outputFile, string zonePivot, string accessToken,
                bool debug, IMarkdownToDocx converter)
        {
            if (string.IsNullOrEmpty(repo))
                throw new ArgumentException($"'{nameof(repo)}' cannot be null or empty.", nameof(repo));
            if (string.IsNullOrEmpty(branch))
                throw new ArgumentException($"'{nameof(branch)}' cannot be null or empty.", nameof(branch));
            if (string.IsNullOrEmpty(folder))
                throw new ArgumentException($"'{nameof(folder)}' cannot be null or empty.", nameof(folder));

            accessToken = string.IsNullOrEmpty(accessToken)
                ? GithubHelper.ReadDefaultSecurityToken()
                : accessToken;

            await Convert(
                TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken),
                accessToken, folder, outputFile, zonePivot, debug, converter);
        }

        public static async Task ConvertFromFolderAsync(string learnFolder, string zonePivot, string outputFile,
                bool debug, IMarkdownToDocx converter)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            await Convert(
                TripleCrownGitHubService.CreateLocal(learnFolder),
                null, learnFolder, outputFile, zonePivot, debug, converter);
        }

        private static async Task Convert(ITripleCrownGitHubService tcService,
            string accessToken, string moduleFolder, string docxFile,
            string zonePivot, bool debug, IMarkdownToDocx converter)
        {
            if (converter is null)
                throw new ArgumentNullException(nameof(converter));

            var rootTemp = debug ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Path.GetTempPath();
            var tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            while (Directory.Exists(tempFolder))
            {
                tempFolder = Path.Combine(rootTemp, Path.GetRandomFileName());
            }

            // Download the module
            var (module, markdownFile) = await new LearnUtilities().DownloadModuleAsync(tcService, accessToken, moduleFolder, tempFolder);

            try
            {
                // Convert the file.
                await converter.Convert(module, markdownFile, docxFile, zonePivot);
            }
            finally
            {
                if (!debug)
                {
                    Directory.Delete(tempFolder);
                }
            }
        }
    }
}