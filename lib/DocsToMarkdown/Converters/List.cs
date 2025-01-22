using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
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
        //prefix = converter.ParentPrefix + prefix;
        converter.PushParent(htmlInput.Name);

        var sb = new StringBuilder();
        var listItems = htmlInput.SelectNodes("li");
        var liConverter = new ListItem();
        
        foreach (var item in listItems)
        {
            var text = liConverter.Convert(converter, item).Trim('\r', '\n', ' ');
            if (text.Contains('\n'))
            {
                //text = Regex.Replace(text, @"[\r\n]{2,}", Environment.NewLine);
                text = text.Replace("\n", "\n" + converter.ParentPrefix);
            }

            if (!string.IsNullOrWhiteSpace(text))
                sb.AppendLine(prefix + text);
        }

        var result = converter.PopParent();
        Debug.Assert(result == htmlInput.Name);
        
        return sb.ToString();
    }

    private bool CheckForHiddenList(HtmlNode htmlInput)
    {
        // Ignore the metadata list. 
        return htmlInput.GetAttributeValue("class", "") == "metadata page-metadata";
    }
}

