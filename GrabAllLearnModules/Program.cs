if (args.Length == 0)
{
    Console.WriteLine("Missing root Learn repo folder to process");
    return;
}

string docPath = Path.Combine(GetDownloadFolderPath(), "learnDocs");
if (!Directory.Exists(docPath))
    Directory.CreateDirectory(docPath);

foreach (var index in Directory.GetFiles(args[0], "index.yml", SearchOption.AllDirectories))
{
    string folder = Path.GetDirectoryName(index);
    if (Directory.GetFiles(folder).Length == 1) 
        continue;

    string docFile = Path.ChangeExtension(Path.GetFileName(folder), "docx");
    string fullDocPath = Path.Combine(docPath, docFile);
    
    if (!File.Exists(fullDocPath))
    {
        if (await ConvertLearnToDoc.Program.Main(new[] {"-n", $"-i{folder}", $"-o{fullDocPath}"}) != 0)
            break;
    }
}

static string GetHomePath()
    => Environment.OSVersion.Platform == PlatformID.Unix 
        ? Environment.GetEnvironmentVariable("HOME") 
        : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");


static string GetDownloadFolderPath() => Path.Combine(GetHomePath(), "Downloads");