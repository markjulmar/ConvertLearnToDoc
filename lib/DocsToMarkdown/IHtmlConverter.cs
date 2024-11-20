using HtmlAgilityPack;
using Julmar.DocsToMarkdown.Metadata;

namespace Julmar.DocsToMarkdown;

public interface IHtmlConverter
{
    public static async Task<IHtmlConverter> LoadAsync(Uri url, HtmlConverterOptions? options = null)
    {
        var htmlDoc = await new HtmlWeb().LoadFromWebAsync(url.ToString());
        return await new HtmlConverter(url, htmlDoc, options).Initialize();
    }

    public Uri Url { get; }
    bool IsTrainingModulePage { get; }
    IDocsMetadata Metadata { get; }
    IEnumerable<string> GetMarkdown();
    IAsyncEnumerable<IHtmlConverter> GetTrainingUnitsConverter();
}