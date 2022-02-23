using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using CompareAll.Comparer;

namespace CompareAll;

public static class Program
{
    private static CommandLineOptions options;

    static async Task Main(string[] args)
    {
        new Parser(cfg => { cfg.HelpWriter = Console.Error; })
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

        string downloadFolder = GetDownloadFolderPath();

        string docPath = Path.Combine(downloadFolder, "learnDocs");
        if (!Directory.Exists(docPath))
            Directory.CreateDirectory(docPath);

        string modulePath = Path.Combine(downloadFolder, "learnModules");
        if (!Directory.Exists(modulePath))
            Directory.CreateDirectory(modulePath);

        foreach (var index in Directory.GetFiles(options.InputFolder, "index.yml", SearchOption.AllDirectories))
        {
            string folder = Path.GetDirectoryName(index);
            if (Directory.GetFiles(folder).Length == 1) // skip learning paths
                continue;

            string fullDocPath = Path.Combine(docPath, Path.ChangeExtension(Path.GetFileName(folder), "docx"));
            string moduleFolder = Path.Combine(modulePath, Path.GetFileName(folder));

            try
            {
                if (await ProcessOneFolder(folder, fullDocPath))
                {
                    if (await ProcessOneFile(fullDocPath, moduleFolder))
                    {
                        await RunCompare(folder, moduleFolder);
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

    private static async Task RunCompare(string originalFolder, string generatedFolder)
    {
        bool outputCsv = options.OutputType == PrintType.Csv;
        string sep = outputCsv ? "," : ": ";
        string lineSep = outputCsv ? string.Empty : new string('-', 30);
        
        string compareOutput = Path.Combine(generatedFolder, outputCsv ? "delta.csv" : "delta.txt");
        if (File.Exists(compareOutput))
            File.Delete(compareOutput);

        await using var sw = new StreamWriter(compareOutput);
        if (outputCsv)
        {
            await sw.WriteLineAsync("Filename,Change,Line,Original,New");
        }

        // YAML files
        foreach (var yamlFile in Directory.GetFiles(originalFolder, "*.yml"))
        {
            string fn = Path.GetFileName(yamlFile);
            string genFile = Path.Combine(generatedFolder, fn);

            bool hasChanges = false;
            if (File.Exists(genFile))
            {
                foreach (var diff in FileComparer.Yaml(yamlFile, genFile))
                {
                    var line = diff.Print(options.OutputType);
                    await sw.WriteLineAsync($"{fn}{sep}{line}");
                    hasChanges = true;
                }

            }
            else
            {
                await sw.WriteLineAsync($"{fn}{sep}missing");
                hasChanges = true;
            }

            if (!outputCsv && hasChanges)
            {
                await sw.WriteLineAsync();
                await sw.WriteLineAsync(lineSep);
                await sw.WriteLineAsync();
            }
        }

        originalFolder = Path.Combine(originalFolder, "includes");
        generatedFolder = Path.Combine(generatedFolder, "includes");

        // Markdown files
        foreach (var mdFile in Directory.GetFiles(originalFolder, "*.md"))
        {
            string fn = Path.GetFileName(mdFile);
            string genFile = Path.Combine(generatedFolder, fn);

            bool hasChanges = false;
            if (File.Exists(genFile))
            {
                foreach (var diff in FileComparer.Markdown(mdFile, genFile))
                {
                    var line = diff.Print(options.OutputType);
                    await sw.WriteLineAsync($"{fn}{sep}{line}");
                    hasChanges = true;
                }

            }
            else
            {
                await sw.WriteLineAsync($"{fn}{sep}missing");
                hasChanges = true;
            }

            if (!outputCsv && hasChanges)
            {
                await sw.WriteLineAsync();
                await sw.WriteLineAsync(lineSep);
                await sw.WriteLineAsync();
            }
        }
    }

    private static async Task<bool> ProcessOneFile(string inputDoc, string outputFolder)
    {
        if (!File.Exists(inputDoc)) throw new ArgumentException($"{inputDoc} does not exist.");
        if (outputFolder == null) throw new ArgumentNullException(nameof(outputFolder));

        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);

            var args = new System.Collections.Generic.List<string> {$"-i{inputDoc}", $"-o{outputFolder}"};
            if (options.Debug) args.Add("-d");

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