using MSLearnRepos;

namespace LearnDocUtils
{
    public static class DocsDownloader
    {
        public static async Task<string> DownloadFileAsync(ILearnRepoService learnRepo, string inputFile, string outputFolder)
        {
            var (_, text) = await learnRepo.ReadFileForPathAsync(inputFile);

            string filename = Path.GetFileName(inputFile);
            string outputFile = Path.Combine(outputFolder, filename);

            await File.WriteAllTextAsync(outputFile, text);
            return outputFile;
        }
    }
}
