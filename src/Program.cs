using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using LearnDocUtils;
using Microsoft.VisualBasic;

namespace ConvertLearnToDoc
{
    public static class Program
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("Learn/Docx converter");

            CommandLineOptions options = null;
            new Parser(cfg => { cfg.HelpWriter = Console.Error; })
                .ParseArguments<CommandLineOptions>(args)
                .WithParsed(clo => options = clo);
            if (options == null)
                return; // bad arguments or help.

            try
            {

                // Input is a Learn module URL
                if (options.InputFileOrFolder.StartsWith("http"))
                {
                    await LearnToDocx.ConvertFromUrl(options.InputFileOrFolder,
                        options.OutputFileOrFolder, options.ZonePivot,
                        options.AccessToken, Console.WriteLine, options.Debug, options.UsePandoc);
                }

                // Input is a repo + folder + branch
                else if (!string.IsNullOrEmpty(options.GitHubRepo))
                {
                    await LearnToDocx.ConvertFromRepo(options.GitHubRepo, options.GitHubBranch,
                        options.InputFileOrFolder,
                        options.OutputFileOrFolder, options.ZonePivot,
                        options.AccessToken, Console.WriteLine, options.Debug, options.UsePandoc);
                }
                // Input is a local folder containing a Learn module
                else if (Directory.Exists(options.InputFileOrFolder))
                {
                    await LearnToDocx.ConvertFromFolder(options.InputFileOrFolder, options.ZonePivot,
                        options.OutputFileOrFolder,
                        Console.WriteLine, options.Debug, options.UsePandoc);
                }
                // Input is a docx file
                else
                {
                    if (string.IsNullOrEmpty(options.OutputFileOrFolder))
                        options.OutputFileOrFolder = Path.ChangeExtension(options.InputFileOrFolder, "");

                    await new DocxToLearn().ConvertAsync(options.InputFileOrFolder, options.OutputFileOrFolder,
                        debug: options.Debug);

                    if (options.ZipOutput)
                    {
                        string baseFolder = Path.GetDirectoryName(options.OutputFileOrFolder);
                        if (string.IsNullOrEmpty(baseFolder))
                            baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                        string zipFile = Path.Combine(baseFolder,
                            Path.ChangeExtension(Path.GetFileNameWithoutExtension(options.OutputFileOrFolder), "zip"));

                        Utils.CompressFolder(options.OutputFileOrFolder, zipFile);
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