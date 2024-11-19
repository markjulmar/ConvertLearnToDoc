using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Emphasis() : BaseConverter("em", "i")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        if (htmlInput.HasChildNodes)
        {
            var result = converter.ConvertChildren(htmlInput);
            return $"_{result}_";
        }

        return $"_{System.Net.WebUtility.HtmlDecode(htmlInput.InnerText)}_";
    }
}