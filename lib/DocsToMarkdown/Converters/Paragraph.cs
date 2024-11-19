using System.Diagnostics;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Paragraph() : BaseConverter("p")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        var result = converter.ConvertChildren(htmlInput);

        if (!string.IsNullOrWhiteSpace(result))
        {
            // Special case returning raw HTML anchor
            if (result.StartsWith("<a ") && result.EndsWith("</a>"))
                return result;
            
            return result + Environment.NewLine;
        }

        return string.Empty;
    }
}