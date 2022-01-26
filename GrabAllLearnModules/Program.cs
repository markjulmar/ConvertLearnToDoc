if (args.Length == 0)
{
    Console.WriteLine("Missing root Learn repo folder to process");
    return;
}

foreach (var index in Directory.GetFiles(args[0], "index.yml", SearchOption.AllDirectories))
{
    string folder = Path.GetDirectoryName(index);
    if (Directory.GetFiles(folder).Length == 1) 
        continue;

    string docFile = Path.ChangeExtension(Path.GetFileName(folder), "docx");

    string fullDocPath = Path.Combine(GetDownloadFolderPath(), "learnDocs", docFile);

    if (!File.Exists(fullDocPath))
    {
        await ConvertLearnToDoc.Program.Main(new[] { $"-i{folder}", $"-o{fullDocPath}" });
        //Console.WriteLine("Press [ENTER] to continue.");
        //Console.ReadLine();
    }
}

static string GetHomePath()
    => Environment.OSVersion.Platform == PlatformID.Unix 
        ? Environment.GetEnvironmentVariable("HOME") 
        : Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");


static string GetDownloadFolderPath() => Path.Combine(GetHomePath(), "Downloads");