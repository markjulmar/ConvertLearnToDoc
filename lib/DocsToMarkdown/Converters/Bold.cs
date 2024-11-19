using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Bold() : BaseConverter("b", "bold", "strong")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        if (htmlInput.HasChildNodes)
        {
            var result = converter.ConvertChildren(htmlInput);
            return $"**{result}**";
        }

        return $"**{System.Net.WebUtility.HtmlDecode(htmlInput.InnerText)}**";
    }
}