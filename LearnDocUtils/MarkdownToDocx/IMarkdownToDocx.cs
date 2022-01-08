using System.Threading.Tasks;
using MSLearnRepos;

namespace LearnDocUtils
{
    public interface IMarkdownToDocx
    {
        Task Convert(TripleCrownModule moduleData,
            string markdownFile, string docxFile, string zonePivot);
    }
}