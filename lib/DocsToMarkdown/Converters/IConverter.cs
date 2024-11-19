using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal interface IConverter
{
    public bool CanConvert(HtmlNode htmlInput);
    public string? Convert(HtmlConverter converter, HtmlNode htmlInput);
}