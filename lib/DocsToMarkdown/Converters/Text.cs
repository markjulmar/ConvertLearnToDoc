using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Text() : BaseConverter("#text")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        return System.Net.WebUtility.HtmlDecode(htmlInput.InnerText);
    }
}