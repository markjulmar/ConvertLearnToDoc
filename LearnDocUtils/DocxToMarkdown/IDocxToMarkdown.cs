using System.Threading.Tasks;

namespace LearnDocUtils
{
    public interface IDocxToMarkdown
    {
        Task ConvertAsync(string docxFile, string markdownFile, string mediaFolder);
    }
}