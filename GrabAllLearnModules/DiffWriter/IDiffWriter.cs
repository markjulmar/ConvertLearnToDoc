using System;
using System.Threading.Tasks;
using FileComparisonLib;

namespace CompareAll.DiffWriter
{
    public interface IDiffWriter : IDisposable
    {
        Task WriteDiffHeaderAsync(string originalFolder, string generatedFolder);
        Task WriteFileHeaderAsync(string filename);
        Task WriteDifferenceAsync(string filename, Difference difference);
        Task WriteMissingFileAsync(string filename);
    }
}
