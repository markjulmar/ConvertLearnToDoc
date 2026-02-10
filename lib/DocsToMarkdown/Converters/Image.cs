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

        string imageText = "";
        
        if (converter.UseDocfxExtensions)
        {
            var type = isComplex ? "type=\"complex\"" : 
                            isIcon ? "type=\"icon\"" :
                            "type=\"content\"";
            var border = hasBorder ? "border=\"true\" " : "";
            var lightbox = lightBoxUrl != null ? $" lightbox=\"{lightBoxUrl}\"" : "";

            var captionText = LookAheadForImageCaptionText(htmlInput);

            var sb = new StringBuilder(
                $":::image {type} {border}source=\"{source}\" alt-text=\"{altText}\"{lightbox}:::");
            if (captionText.Length > 0)
            {
                sb.AppendLine(captionText);
                sb.Append(":::image-end:::");
            }

            imageText = sb.ToString();
        }

        else if (lightBoxUrl != null)
        {
            imageText = new StringBuilder("<a href=\"")
                .Append(lightBoxUrl)
                .Append("\">")
                .Append("![")
                .Append(altText)
                .Append("](")
                .Append(source)
                .Append(")</a>")
                .ToString();            
        }
        
        else imageText = $"![{altText}]({source})";
        
        return converter.ParentPrefix != "" ? Environment.NewLine + imageText : imageText;
    }

    private static string LookAheadForImageCaptionText(HtmlNode htmlInput)
    {
        // <p><img src="test.png" alt="test"></p>
        // <div class="visually-hidden"><p>
        // Some Graphic.</p></div>
        var testNode = htmlInput.ParentNode.NextSibling;
        return testNode?.HasClass("visually-hidden") == true
            ? testNode.InnerText.Replace("\n", Environment.NewLine)
            : string.Empty;
    }
}