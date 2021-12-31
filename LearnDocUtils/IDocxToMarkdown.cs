using System;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public interface IDocxToMarkdown
    {
        Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder, Action<string> logger = null, bool debug = false);
        Func<string,string> MarkdownProcessor { get; }
    }
}