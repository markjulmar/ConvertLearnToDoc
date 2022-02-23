using System.IO;
using System.Threading.Tasks;
using CompareAll.Comparer;

namespace CompareAll.DiffWriter;

public class TextDiffWriter : IDiffWriter
{
    private static readonly string FileSeparator = new('-', 30);
    private static string separator = "  ";
    private readonly TextWriter writer;

    public TextDiffWriter(string filename)
    {
        this.writer = new StreamWriter(Path.ChangeExtension(filename, "txt"));
    }

    public async Task WriteDiffHeaderAsync(string originalFolder, string generatedFolder)
    {
        await writer.WriteLineAsync($"Diff compare of {originalFolder} to {generatedFolder}");
        await writer.WriteLineAsync();
    }

    public async Task WriteFileHeaderAsync(string filename)
    {
        await writer.WriteLineAsync(FileSeparator);
        await writer.WriteLineAsync(filename);

    }

    public async Task WriteDifferenceAsync(string filename, Difference difference)
    {
        string line = difference.ToString()
            .Trim('\r', '\n')
            .Replace("\n", $"\n{separator}");

        await writer.WriteLineAsync($"{separator}{line}");
    }

    public async Task WriteMissingFileAsync(string filename)
    {
        await writer.WriteLineAsync(FileSeparator);
        await writer.WriteLineAsync($"{filename} missing.");
    }

    public void Dispose() => writer?.Dispose();
}