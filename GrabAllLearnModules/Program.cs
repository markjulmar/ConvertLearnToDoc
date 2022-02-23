using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommandLine;
using CompareAll.Comparer;
using CompareAll.DiffWriter;

namespace CompareAll;

public static class Program
{
    public static async Task Main(string[] args)
    {
        CommandLineOptions options = null;
        new Parser(cfg =>
            {
                cfg.HelpWriter = Console.Error;
                cfg.CaseInsensitiveEnumValues = true;
            })
            .ParseArguments<CommandLineOptions>(args)
            .WithParsed(clo => options = clo);
        if (options == null)
            return; // bad arguments or help.

        Console.WriteLine($"CompareAll: testing folder {Path.GetFileName(options.InputFolder)}");

        if (!Directory.Exists(options.InputFolder))
        {
            Console.WriteLine($"{options.InputFolder} does not exist.");
            return;
        }

        // All looks good.
        await Run(options);
    }

    private static async Task Run(CommandLineOptions options)
    {
        string outputFolder = options.OutputFolder ?? GetDownloadFolderPath();

        string docPath = Path.Combine(outputFolder, "learnDocs");
        if (!Directory.Exists(docPath))
            Directory.CreateDirectory(docPath);

        string modulePath = Path.Combine(outputFolder, "learnModules");
        if (!Directory.Exists(modulePath))
            Directory.CreateDirectory(modulePath);

        foreach (var index in Directory.GetFiles(options.InputFolder, "index.yml", SearchOption.AllDirectories))
        {
            string folder = Path.GetDirectoryName(index) ?? "";
            if (Directory.GetFiles(folder).Length < 2) // skip learning paths
                continue;

            string fullDocPath = Path.Combine(docPath, Path.ChangeExtension(Path.GetFileName(folder), "docx"));
            string moduleFolder = Path.Combine(modulePath, Path.GetFileName(folder));

            try
            {
                if (await ProcessOneFolder(folder, fullDocPath))
                {
                    if (await ProcessOneFile(fullDocPath, moduleFolder, options.Debug))
                    {
                        await RunCompare(folder, moduleFolder, options.OutputType);
                    }
                }
            }
            catch
            {
                File.Delete(fullDocPath);
                Directory.Delete(moduleFolder, true);
                throw;
            }
        }
    }

    private static async Task RunCompare(string originalFolder, string generatedFolder, PrintType outputType)
    {
        string diffFile = Path.Combine(generatedFolder, "diff.tmp");
        using IDiffWriter diffWriter = outputType switch
        {
            PrintType.Csv => new CsvDiffWriter(diffFile),
            PrintType.Markdown => new MarkdownDiffWriter(diffFile),
            PrintType.Text => new TextDiffWriter(diffFile),
            _ => new TextDiffWriter(diffFile)
        };

        // Write the header
        await diffWriter.WriteDiffHeaderAsync(originalFolder, generatedFolder);

        // YAML files
        await ProcessDiffs(originalFolder, generatedFolder, "*.yaml", diffWriter, FileComparer.Yaml);

        // Markdown files
        await ProcessDiffs(Path.Combine(originalFolder, "includes"), Path.Combine(generatedFolder, "includes"),
            "*.md", diffWriter, FileComparer.Markdown);
    }

    private static async Task ProcessDiffs(string originalFolder, string generatedFolder, string fileSpec,
        IDiffWriter writer, Func<string, string, IEnumerable<Difference>> diffProcessor)
    {
        foreach (var yamlFile in Directory.GetFiles(originalFolder, fileSpec))
        {
            string filename = Path.GetFileName(yamlFile);
            string genFile = Path.Combine(generatedFolder, filename);

            if (File.Exists(genFile))
            {
                var lines = diffProcessor.Invoke(yamlFile, genFile).ToList();
                if (lines.Count == 0)
                    continue;

                // Write the header
                await writer.WriteFileHeaderAsync(filename);

                // Write each diff line
                foreach (var difference in lines)
                {
                    await writer.WriteDifferenceAsync(filename, difference);
                }
            }
            else
            {
                await writer.WriteMissingFileAsync(filename);
            }
        }
    }

    private static async Task<bool> ProcessOneFile(string inputDoc, string outputFolder, bool debug)
    {
        if (!File.Exists(inputDoc)) throw new ArgumentException($"{inputDoc} does not exist.");
        if (outputFolder == null) throw new ArgumentNullException(nameof(outputFolder));

        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);

            var args = new List<string> {$"-i{inputDoc}", $"-o{outputFolder}"};
            if (debug) args.Add("-d");

            if (await ConvertLearnToDoc.Program.Main(args.ToArray()) != 0)
                return false;
        }
        return true;
    }

    private static async Task<bool> ProcessOneFolder(string inputFolder, string outputDoc)
    {
        if (inputFolder == null) throw new ArgumentNullException(nameof(inputFolder));
        if (outputDoc == null) throw new ArgumentNullException(nameof(outputDoc));
        if (!Directory.Exists(inputFolder)) throw new ArgumentException($"{inputFolder} does not exist.");

        if (!File.Exists(outputDoc))
        {
            if (await ConvertLearnToDoc.Program.Main(new[] { $"-i{inputFolder}", $"-o{outputDoc}" }) != 0)
                return false;
        }

        return true;
    }

    private static string GetHomePath()
        => Environment.OSVersion.Platform == PlatformID.Unix
            ? Environment.GetEnvironmentVariable("HOME")
            : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");


    private static string GetDownloadFolderPath() => Path.Combine(GetHomePath(), "Downloads");
}