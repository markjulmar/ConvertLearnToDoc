using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class List() : BaseConverter("ul", "ol")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));

        if (CheckForHiddenList(htmlInput))
            return string.Empty;
        
        var prefix = htmlInput.Name.Equals("ol", StringComparison.InvariantCultureIgnoreCase)
            ? "1. "
            : "- ";

        var sb = new StringBuilder();
        var listItems = htmlInput.SelectNodes("li");
        var liConverter = new ListItem();
        
        foreach (var item in listItems)
        {
            var text = liConverter.Convert(converter, item).Trim('\r', '\n', ' ');
            if (!string.IsNullOrWhiteSpace(text))
                sb.AppendLine(prefix + text);
        }

        return sb.ToString();
    }

    private bool CheckForHiddenList(HtmlNode htmlInput)
    {
        // Ignore the metadata list. 
        return htmlInput.GetAttributeValue("class", "") == "metadata page-metadata";
    }
}

