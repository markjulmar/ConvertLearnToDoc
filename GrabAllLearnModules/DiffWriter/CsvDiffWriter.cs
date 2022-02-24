using System.IO;
using System.Threading.Tasks;
using FileComparisonLib;

namespace CompareAll.DiffWriter;

public class CsvDiffWriter : IDiffWriter
{
    private readonly TextWriter writer;

    public CsvDiffWriter(string filename)
    {
        this.writer = new StreamWriter(Path.ChangeExtension(filename, "csv"));
    }

    public async Task WriteDiffHeaderAsync(string originalFolder, string generatedFolder)
    {
        await writer.WriteLineAsync("Filename,Change,Line,Original,New");
    }

    public Task WriteFileHeaderAsync(string filename) => Task.CompletedTask;

    public async Task WriteDifferenceAsync(string filename, Difference difference)
    {
        await writer.WriteLineAsync($"{filename},{difference.Change},{difference.Key},{difference.EscapedOriginalValue},{difference.EscapedNewValue}");
    }

    public async Task WriteMissingFileAsync(string filename)
    {
        await writer.WriteLineAsync($"{filename},missing,,,");
    }

    public void Dispose() => writer?.Dispose();
}