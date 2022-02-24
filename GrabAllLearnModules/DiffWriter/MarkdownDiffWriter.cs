using System.IO;
using System.Threading.Tasks;
using FileComparisonLib;

namespace CompareAll.DiffWriter;

public class MarkdownDiffWriter : IDiffWriter
{
    private readonly TextWriter writer;

    public MarkdownDiffWriter(string filename)
    {
        this.writer = new StreamWriter(Path.ChangeExtension(filename, "md"));
    }

    public async Task WriteDiffHeaderAsync(string originalFolder, string generatedFolder)
    {
        await writer.WriteLineAsync($"# Compare of {originalFolder} to {generatedFolder}");
        await writer.WriteLineAsync();
    }

    public async Task WriteFileHeaderAsync(string filename)
    {
        await writer.WriteLineAsync($"## {filename}");
        await writer.WriteLineAsync();
        await writer.WriteLineAsync("| Type | Position | Original | New |");
        await writer.WriteLineAsync("|------|----------|----------|-----|");
    }

    public async Task WriteDifferenceAsync(string filename, Difference difference)
    {
        await writer.WriteLineAsync(
            $"| **{difference.Change}** | `{difference.Key}` | {difference.EscapedOriginalValue} | {difference.EscapedNewValue} |");
    }

    public async Task WriteMissingFileAsync(string filename)
    {
        await writer.WriteLineAsync($"| **ERROR** {filename} missing. | | | |");
    }

    public void Dispose() => writer?.Dispose();
}