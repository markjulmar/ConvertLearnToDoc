﻿using System;
using System.IO;
using System.Linq;
using CommandLine;
using ConvertLearnToDoc;
using LearnDocUtils;

Console.WriteLine("Learn/Docx converter");

CommandLineOptions options = null;
new Parser(cfg => { cfg.HelpWriter = Console.Error; })
    .ParseArguments<CommandLineOptions>(args)
    .WithParsed(clo => options = clo);
if (options == null)
    return; // bad arguments or help.

if (options.InputFileOrFolder.StartsWith("http"))
{
    var (repo, branch, folder) = await Utils.RetrieveLearnLocationFromUrlAsync(options.InputFileOrFolder);

    if (string.IsNullOrEmpty(options.OutputFileOrFolder))
        options.OutputFileOrFolder = Path.ChangeExtension(folder.Split('/').Last(), "docx");

    await LearnToDocx.ConvertAsync(repo, branch, folder, options.OutputFileOrFolder);
}
else if (Directory.Exists(options.InputFileOrFolder))
{
    if (string.IsNullOrEmpty(options.OutputFileOrFolder))
        options.OutputFileOrFolder = Path.ChangeExtension(options.InputFileOrFolder, "docx");

    await LearnToDocx.ConvertAsync(options.InputFileOrFolder, options.OutputFileOrFolder);
}
else
{
    if (string.IsNullOrEmpty(options.OutputFileOrFolder))
        options.OutputFileOrFolder = Path.ChangeExtension(options.InputFileOrFolder, "");
    
    await DocxToLearn.ConvertAsync(options.InputFileOrFolder, options.OutputFileOrFolder);

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