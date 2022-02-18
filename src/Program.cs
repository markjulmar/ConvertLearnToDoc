using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using LearnDocUtils;

namespace ConvertLearnToDoc
{
    public static class Program
    {
        public static async Task<int> Main(string[] args)
        {
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
                    var log = await LearnToDocx.ConvertFromUrlAsync(options.InputFileOrFolder,
                        options.OutputFileOrFolder, options.ZonePivot, options.AccessToken, new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks});

                    log.ForEach(Console.WriteLine);
                }

                // Input is a repo + folder + branch
                else if (!string.IsNullOrEmpty(options.GitHubRepo))
                {
                    await LearnToDocx.ConvertFromRepoAsync(options.GitHubRepo, options.GitHubBranch,
                        options.InputFileOrFolder, options.OutputFileOrFolder, options.ZonePivot, options.AccessToken, new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks });

                }
                // Input is a local folder containing a Learn module
                else if (Directory.Exists(options.InputFileOrFolder))
                {
                    await LearnToDocx.ConvertFromFolderAsync(options.InputFileOrFolder, options.OutputFileOrFolder, options.ZonePivot, new DocumentOptions { Debug = options.Debug, EmbedNotebookContent = options.ConvertNotebooks });
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
        }
    }
}