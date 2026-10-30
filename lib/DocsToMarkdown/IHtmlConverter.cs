using HtmlAgilityPack;
using Julmar.DocsToMarkdown.Metadata;

namespace Julmar.DocsToMarkdown;

public interface IHtmlConverter
{
    public static async Task<IHtmlConverter> LoadAsync(Uri url, HtmlConverterOptions? options = null)
    {
        using var client = new HttpClient();
        using var response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(await response.Content.ReadAsStringAsync());

        // Relative asset paths belong to the page after any redirects, not the
        // original request URL (which may have an extra trailing slash).
        var pageUrl = response.RequestMessage?.RequestUri ?? url;
        return await new HtmlConverter(pageUrl, htmlDoc, options).Initialize();
    }

    public Uri Url { get; }
    bool IsTrainingModulePage { get; }
    IDocsMetadata Metadata { get; }
    IEnumerable<string> GetMarkdown();
    IAsyncEnumerable<IHtmlConverter> GetTrainingUnitsConverter();
    string ParentPrefix { get; }
    void PushParent(string parent);
    string? PopParent();
}
