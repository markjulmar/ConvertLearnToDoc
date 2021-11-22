using System;
using System.IO;
using System.Linq;
using LearnDocUtils;

Console.WriteLine("Learn/Docx converter");

if (args.Length < 1 || args.Length > 2)
{
    Console.WriteLine("Missing: [input] [output]");
    Console.WriteLine("Where input/output can be a Learn source folder, or Word doc.");
    return;
}

string inputFile = args[0];
string outputFile;

if (inputFile.StartsWith("http"))
{
    (string repo, string branch, string folder) = await Utils.RetrieveLearnLocationFromUrlAsync(inputFile);

    if (args.Length == 1)
    {
        outputFile = Path.ChangeExtension(folder.Split('/').Last(), "docx");
        await LearnToDocx.ConvertAsync(repo, branch, folder, outputFile);

    }
    else outputFile = args[1];
}
else if (Directory.Exists(inputFile))
{
    outputFile = args.Length == 1 ? Path.ChangeExtension(args[0], "docx") : args[1];
    await LearnToDocx.ConvertAsync(inputFile, outputFile);
}
else
{
    outputFile = args.Length == 1 ? Path.ChangeExtension(args[0], "") : args[1];
    await DocxToLearn.ConvertAsync(inputFile, outputFile);
}

Console.WriteLine($"Converted {inputFile} to {outputFile}.");