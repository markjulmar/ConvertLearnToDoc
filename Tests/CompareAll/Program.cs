using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using CommandLine;
using CompareAll.DiffWriter;
using FileComparisonLib;
using LearnDocUtils;
using FileComparer = FileComparisonLib.FileComparer;
using CompareAll;

CommandLineOptions options = null;
new Parser(cfg => { cfg.HelpWriter = Console.Error; cfg.CaseInsensitiveEnumValues = true; })
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

    Console.WriteLine($"Processing {Path.GetFileName(folder)}");

    try
    {
        await ProcessOneModuleAsync(folder, fullDocPath);
        await ProcessOneWordDocAsync(fullDocPath, moduleFolder, options.Debug);
        await RunCompareAsync(folder, moduleFolder, options.OutputType);
    }
    catch
    {
        File.Delete(fullDocPath);
        Directory.Delete(moduleFolder, true);
        throw;
    }
}

static async Task RunCompareAsync(string originalFolder, string generatedFolder, PrintType outputType)
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
    await ProcessDiffsAsync(originalFolder, generatedFolder, "*.yaml", diffWriter, FileComparer.Yaml);

    // Markdown files
    await ProcessDiffsAsync(Path.Combine(originalFolder, "includes"), Path.Combine(generatedFolder, "includes"),
        "*.md", diffWriter, FileComparer.Markdown);
}

static async Task ProcessDiffsAsync(string originalFolder, string generatedFolder, string fileSpec,
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

static async Task ProcessOneWordDocAsync(string inputDoc, string outputFolder, bool debug)
{
    if (!File.Exists(inputDoc)) throw new ArgumentException($"{inputDoc} does not exist.");
    if (outputFolder == null) throw new ArgumentNullException(nameof(outputFolder));

    await DocxToLearn.ConvertAsync(inputDoc, outputFolder, new MarkdownOptions { Debug = debug });
}

static async Task ProcessOneModuleAsync(string inputFolder, string outputDoc)
{
    if (inputFolder == null) throw new ArgumentNullException(nameof(inputFolder));
    if (outputDoc == null) throw new ArgumentNullException(nameof(outputDoc));
    if (!Directory.Exists(inputFolder)) throw new ArgumentException($"{inputFolder} does not exist.");

    if (File.Exists(outputDoc))
        File.Delete(outputDoc);

    await LearnToDocx.ConvertFromFolderAsync(inputFolder, outputDoc);
}

static string GetHomePath()
    => Environment.OSVersion.Platform == PlatformID.Unix
        ? Environment.GetEnvironmentVariable("HOME")
        : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");


static string GetDownloadFolderPath() => Path.Combine(GetHomePath(), "Downloads");
