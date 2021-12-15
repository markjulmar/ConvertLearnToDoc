using System;
using System.Threading.Tasks;
using MSLearnRepos;

namespace LearnDocUtils
{
    public interface ILearnToDocx
    {
        Task Convert(ITripleCrownGitHubService tcService, string accessToken,
            string moduleFolder, string outputFile, string zonePivot,
            Action<string> logger, bool debug);
    }
}