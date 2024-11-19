using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Anchor() : BaseConverter("a")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        // TODO -- look to data tags
        
        var url = htmlInput.GetAttributeValue("href", "");
        if (string.IsNullOrEmpty(url)) // emit the tag directly if no URL.
            return htmlInput.OuterHtml;

        // Handle lightbox images.
        if (url.EndsWith("#lightbox", StringComparison.InvariantCultureIgnoreCase)
            && htmlInput.ChildNodes.Any(cn => cn.Name.Equals("img", StringComparison.CurrentCultureIgnoreCase)))
        {
            return new Image().Convert(converter, htmlInput.SelectSingleNode("img"));
        }

        // More than text and not an image, just emit the tag.
        if (htmlInput.ChildNodes.Count > 1)
            return htmlInput.OuterHtml;
        
        // Add MD to end
        if (!url.Contains('/') && !url.Contains('.') && !url.Contains('#'))
            url = url.Trim() + ".md";

        // Grab text -- use the URL if missing.
        var anchorText = System.Net.WebUtility.HtmlDecode(htmlInput.InnerText);
        if (string.IsNullOrWhiteSpace(anchorText))
            anchorText = url;
        
        return $"[{anchorText}]({url})";
    }
}