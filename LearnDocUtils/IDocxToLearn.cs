using System;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public interface IDocxToLearn
    {
        Task ConvertAsync(string docxFile, string outputFolder, Action<string> logger = null, bool debug = false);
    }
}