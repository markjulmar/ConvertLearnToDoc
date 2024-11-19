using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class ListItem() : BaseConverter("li")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        return converter.ConvertChildren(htmlInput) ?? string.Empty;
    }
}