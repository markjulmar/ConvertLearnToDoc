using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Image() : BaseConverter("img", "image")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));

        var altText = htmlInput.Attributes["alt"].Value;
        var originalSource = htmlInput.Attributes["src"].Value;
        var source = SimplifyRelativePath(converter, originalSource);

        // Attempt to download.
        converter.DownloadAsset(source, new Uri(converter.Url, originalSource));
        
        string? lightBoxUrl = null;
        var parentNode = htmlInput.ParentNode;
        if (parentNode.Name.Equals("a", StringComparison.InvariantCultureIgnoreCase))
        {
            lightBoxUrl = parentNode.Attributes["href"].Value;
            if (lightBoxUrl.EndsWith("#lightbox", StringComparison.InvariantCultureIgnoreCase))
                lightBoxUrl = lightBoxUrl[..^9];
            parentNode = parentNode.ParentNode;
        }

        bool hasBorder = (parentNode.Attributes.Contains("class") 
                          && parentNode.Attributes["class"].Value == "mx-imgBorder");

        bool isIcon = htmlInput.Attributes.Contains("presentation") 
            && htmlInput.Attributes["presentation"].Value == "presentation";

        bool isComplex = htmlInput.Attributes.Contains("aria-describedby"); 
        
        if (converter.UseDocfxExtensions)
        {
            var type = isComplex ? "type=\"complex\"" : 
                            isIcon ? "type=\"icon\"" :
                            "type=\"content\"";
            var border = hasBorder ? "border=\"true\" " : "";
            var lightbox = lightBoxUrl != null ? $" lightbox=\"{lightBoxUrl}\"" : "";
            return $":::image {type} {border}source=\"{source}\" alt-text=\"{altText}\"{lightbox}:::";
        }

        if (lightBoxUrl != null)
        {
            return new StringBuilder("<a href=\"")
                .Append(lightBoxUrl)
                .Append("\">")
                .Append("![")
                .Append(altText)
                .Append("](")
                .Append(source)
                .Append(")</a>")
                .ToString();            
        }
        
        return $"![{altText}]({source})";
    }
}