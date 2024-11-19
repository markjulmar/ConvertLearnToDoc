using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Table(): BaseConverter("table")
{
    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));
        var result = new StringBuilder();

        result.Append(ProcessHead(converter, htmlInput)); // already has CR/LF.
        result.Append(ProcessBody(converter, htmlInput));

        return result.ToString();
    }

    private static string ProcessBody(HtmlConverter converter, HtmlNode htmlInput)
    {
        var bodyNode = htmlInput.SelectSingleNode("tbody");
        if (bodyNode == null)
        {
            Debug.Assert(false);
        }
        
        var result = new StringBuilder();
        foreach (var row in bodyNode.SelectNodes("tr"))
        {
            result.Append('|');
            foreach (var cell in row.SelectNodes("td"))
            {
                var text = converter.ConvertChildren(cell);
                if (string.IsNullOrWhiteSpace(text))
                    text = " ";
                result.Append($"{text}|");
            }
            result.AppendLine();
        }

        return result.ToString();
    }

    private static string ProcessHead(HtmlConverter converter, HtmlNode htmlInput)
    {
        var headerNode = htmlInput.SelectSingleNode("thead");
        if (headerNode == null)
        {
            Debug.Assert(false);
        }
        
        var result = new StringBuilder();
        var separator = new StringBuilder();
        var row = headerNode.SelectSingleNode("tr");
        result.Append('|');
        separator.Append('|');
        
        foreach (var cell in row.SelectNodes("th"))
        {
            var styles = cell.GetAttributeValue("style", "").Split(';');
            var text = converter.ConvertChildren(cell);
            if (string.IsNullOrWhiteSpace(text))
                text = " ";
            result.Append($"{text}|");

            var align = styles.FirstOrDefault(s => s.StartsWith("text-align:", StringComparison.InvariantCultureIgnoreCase));
            if (align != null)
            {
                if (align.EndsWith("center", StringComparison.InvariantCultureIgnoreCase))
                    separator.Append(':' + new string('-', text.Length-2) + ':' + '|');
                else if (align.EndsWith("right", StringComparison.InvariantCultureIgnoreCase))
                    separator.Append(new string('-', text.Length-1) + ':' + '|');
                else
                    separator.Append(':' + new string('-', text.Length-1) + '|');
            }
            else
            {
                separator.Append(new string('-', text.Length) + '|');
            }
        }
        result.AppendLine();
        result.AppendLine(separator.ToString());

        return result.ToString();
    }
}