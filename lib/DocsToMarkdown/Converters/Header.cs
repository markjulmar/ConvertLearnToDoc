using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Header() : BaseConverter("h1", "h2", "h3", "h4", "h5", "h6")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        int headerLevel = base.SupportedTags.IndexOf(htmlInput.Name.ToLower());
        Debug.Assert(headerLevel >= 0);
        return $"{new string('#', headerLevel + 1)} {System.Net.WebUtility.HtmlDecode(htmlInput.InnerText)}{Environment.NewLine}";
    }
}