using System;
using MSLearnRepos;
using System.IO;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class LearnToDocx
    {
        public static async Task ConvertFromUrl(string url, string outputFile, string zonePivot, string accessToken,
                                            Action<string> logger, bool debug, bool usePanDoc)
        {
            if (string.IsNullOrEmpty(url))
                throw new ArgumentNullException(nameof(url));

            var (repo, branch, folder) = await LearnUtilities.RetrieveLearnLocationFromUrlAsync(url);
            await ConvertFromRepo(repo, branch, folder, outputFile, zonePivot, accessToken, logger, debug, usePanDoc);
        }

        public static async Task ConvertFromRepo(string repo, string branch, string folder,
                                            string outputFile, string zonePivot, string accessToken, 
                                            Action<string> logger, bool debug, bool usePanDoc)
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

            if (usePanDoc)
            {
                await new LearnToDocxPandoc().Convert(
                    TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken), 
                    accessToken, folder, outputFile, null, logger, debug);
            }
            else
            {
                await new LearnToDocxDXPlus().Convert(
                    TripleCrownGitHubService.CreateFromToken(repo, branch, accessToken),
                    accessToken, folder, outputFile, zonePivot, logger, debug);
            }
        }

        public static async Task ConvertFromFolder(string learnFolder, string zonePivot, string outputFile, Action<string> logger, bool debug, bool usePanDoc)
        {
            if (string.IsNullOrWhiteSpace(learnFolder))
                throw new ArgumentException($"'{nameof(learnFolder)}' cannot be null or whitespace.", nameof(learnFolder));

            if (!Directory.Exists(learnFolder))
                throw new DirectoryNotFoundException($"{learnFolder} does not exist.");

            if (usePanDoc)
            {
                await new LearnToDocxPandoc().Convert(
                    TripleCrownGitHubService.CreateLocal(learnFolder),
                    null, learnFolder, outputFile, zonePivot, logger, debug);
            }
            else
            {
                await new LearnToDocxDXPlus().Convert(
                    TripleCrownGitHubService.CreateLocal(learnFolder),
                    null, learnFolder, outputFile, zonePivot, logger, debug);
            }
        }
    }
}