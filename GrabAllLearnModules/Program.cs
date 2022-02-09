using System.Diagnostics;

if (args.Length == 0)
{
    Console.WriteLine("Missing root Learn repo folder to process");
    return;
}

string downloadFolder = GetDownloadFolderPath();

string docPath = Path.Combine(downloadFolder, "learnDocs");
if (!Directory.Exists(docPath))
    Directory.CreateDirectory(docPath);

string modulePath = Path.Combine(downloadFolder, "learnModules");
if (!Directory.Exists(modulePath))
    Directory.CreateDirectory(modulePath);

foreach (var index in Directory.GetFiles(args[0], "index.yml", SearchOption.AllDirectories))
{
    string folder = Path.GetDirectoryName(index);
    if (Directory.GetFiles(folder).Length == 1) // skip learning paths
        continue;

    string fullDocPath = Path.Combine(docPath, Path.ChangeExtension(Path.GetFileName(folder), "docx"));
    if (await ProcessOneFolder(folder, fullDocPath))
    {
        string moduleFolder = Path.Combine(modulePath, Path.GetFileName(folder));
        if (await ProcessOneFile(fullDocPath, moduleFolder))
        {
            await RunCompare(folder, moduleFolder);
            Console.WriteLine("Press ENTER to continue to next file.");
            Console.ReadLine();
        }
    }
}

static async Task RunCompare(string original, string generated)
{
    string doneFn = Path.Combine(generated, "checked.txt");
    if (File.Exists(doneFn)) return;

    const string bc = @"C:\Users\mark\OneDrive\Tools\Beyond Compare 3\bcompare.exe";
    var proc = Process.Start(bc, $"\"{original}\" \"{generated}\"");
    if (proc != null)
    {
        await proc.WaitForExitAsync();
        File.WriteAllText(doneFn,"");
    }
}

static async Task<bool> ProcessOneFile(string doc, string folder)
{
    if (!Directory.Exists(folder))
    {
        Directory.CreateDirectory(folder);

        try
        {
            if (await ConvertLearnToDoc.Program.Main(new[] {"-d", $"-i{doc}", $"-o{folder}"}) != 0)
                return false;
        }
        catch (Exception ex)
        {
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(folder), $"{folder}.txt"),
                $"File: {doc}{Environment.NewLine}" + $"Error:{Environment.NewLine}" + ex);
        }
    }
    return true;
}

static async Task<bool> ProcessOneFolder(string folder, string doc)
{
    if (!File.Exists(doc))
    {
        try
        {
            if (await ConvertLearnToDoc.Program.Main(new[] {"-n", $"-i{folder}", $"-o{doc}"}) != 0)
                return false;
        }
        catch (Exception ex)
        {
            File.WriteAllText(Path.ChangeExtension(doc, "txt"),
                $"Folder: {folder}{Environment.NewLine}" + $"Error:{Environment.NewLine}" + ex);
            return false;
        }
    }
    return true;
}

static string GetHomePath()
    => Environment.OSVersion.Platform == PlatformID.Unix 
        ? Environment.GetEnvironmentVariable("HOME") 
        : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");


static string GetDownloadFolderPath() => Path.Combine(GetHomePath(), "Downloads");