using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using LearnDocUtils;

namespace ConvertLearnToDoc
{
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

            Console.WriteLine("Learn/Docx converter");

            try
            {
                // Input is a Learn module URL
                if (options.InputFileOrFolder.StartsWith("http"))
                {
                    await LearnToDocx.ConvertFromUrlAsync(options.InputFileOrFolder,
                        options.OutputFileOrFolder, options.ZonePivot,
                        options.AccessToken, options.Debug,
                        options.UsePandoc
                            ? MarkdownConverterFactory.WithPandoc
                            : MarkdownConverterFactory.WithDxPlus);
                }

                // Input is a repo + folder + branch
                else if (!string.IsNullOrEmpty(options.GitHubRepo))
                {
                    await LearnToDocx.ConvertFromRepoAsync(options.GitHubRepo, options.GitHubBranch,
                        options.InputFileOrFolder,
                        options.OutputFileOrFolder, options.ZonePivot,
                        options.AccessToken, options.Debug,
                        options.UsePandoc
                            ? MarkdownConverterFactory.WithPandoc
                            : MarkdownConverterFactory.WithDxPlus);

                }
                // Input is a local folder containing a Learn module
                else if (Directory.Exists(options.InputFileOrFolder))
                {
                    await LearnToDocx.ConvertFromFolderAsync(options.InputFileOrFolder, options.ZonePivot,
                        options.OutputFileOrFolder, options.Debug,
                        options.UsePandoc
                            ? MarkdownConverterFactory.WithPandoc
                            : MarkdownConverterFactory.WithDxPlus);
                }
                // Input is a docx file
                else
                {
                    if (string.IsNullOrEmpty(options.OutputFileOrFolder))
                        options.OutputFileOrFolder = Path.ChangeExtension(options.InputFileOrFolder, "");

                    await DocxToLearn.ConvertAsync(options.InputFileOrFolder, options.OutputFileOrFolder,
                        options.Debug, options.UsePandoc
                            ? DocxConverterFactory.WithPandoc
                            : DocxConverterFactory.WithDxPlus);

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

                Console.WriteLine($"Converted {options.InputFileOrFolder} to {options.OutputFileOrFolder}.");
            }
            catch (AggregateException aex)
            {
                if (options.Debug) throw;
                Console.WriteLine(aex.Flatten().Message);
            }
            catch (Exception ex)
            {
                if (options.Debug) throw;
                Console.WriteLine(ex.Message);
            }
        }
    }
}