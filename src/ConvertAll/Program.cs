using CommandLine;
using ConvertAll;
using LearnDocUtils;

CommandLineOptions options = null;
new Parser(cfg => { cfg.HelpWriter = Console.Error; cfg.CaseInsensitiveEnumValues = true; })
    .ParseArguments<CommandLineOptions>(args)
    .WithParsed(clo => options = clo);
if (options == null)
    return; // bad arguments or help.

Console.WriteLine($"ConvertAll: using folder {Path.GetFileName(options.InputFolder)}");

if (!Directory.Exists(options.InputFolder))
{
    Console.WriteLine($"{options.InputFolder} does not exist.");
    return;
}

if (!Directory.Exists(options.OutputFolder))
    Directory.CreateDirectory(options.OutputFolder);

foreach (var index in Directory.GetFiles(options.InputFolder, "index.yml", SearchOption.AllDirectories))
{
    string folder = Path.GetDirectoryName(index) ?? "";
    if (Directory.GetFiles(folder).Length < 2) // skip learning paths
        continue;

    string filename = Path.GetFileName(folder);
    Console.WriteLine($"Processing {filename}");

    string fullDocPath = Path.Combine(options.OutputFolder, Path.ChangeExtension(filename, "docx"));
    if (File.Exists(fullDocPath))
        File.Delete(fullDocPath);

    var results = await LearnToDocx.ConvertFromFolderAsync(folder, fullDocPath, new DocumentOptions { Debug = options.Debug });
    File.WriteAllText(Path.ChangeExtension(fullDocPath, ".txt"), string.Join(Environment.NewLine, results));
}