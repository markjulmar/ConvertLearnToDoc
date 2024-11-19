using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown;

public static class HtmlConverterFactory
{
    public static async Task<IHtmlConverter> Create(Uri url)
    {
        var htmlDoc = await new HtmlWeb().LoadFromWebAsync(url.ToString());
        return await new HtmlConverter(htmlDoc).Initialize();
    }
}