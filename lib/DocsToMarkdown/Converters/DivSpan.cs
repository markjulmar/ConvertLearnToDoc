using System.Diagnostics;
using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class DivSpan() : BaseConverter("div", "span")
{
    readonly Dictionary<string, Func<string,HtmlConverter,HtmlNode,string>> _blockConverters = new()
    {
        { "NOTE", QuoteBlock },
        { "TIP", QuoteBlock },
        { "WARNING", QuoteBlock },
        { "IMPORTANT", QuoteBlock },
        { "ERROR", QuoteBlock },
        { "CAUTION", QuoteBlock },
        { "zone", ZonePivotSection },
        { "has-pivot", ZonePivotSection},
        { "xp-tag-hexagon", IgnoreBlock },
    };

    private static string IgnoreBlock(string arg1, HtmlConverter arg2, HtmlNode arg3)
    {
        // Ignore this block.
        return nameof(IgnoreBlock);
    }

    private static string ZonePivotSection(string className, HtmlConverter converter, HtmlNode htmlNode)
    {
        var pivotName = htmlNode.GetAttributeValue("data-pivot", "");
        if (string.IsNullOrWhiteSpace(pivotName))
        {
            return converter.ConvertChildren(htmlNode) ?? string.Empty;
        }

        return new StringBuilder($"::: zone pivot=\"{pivotName}\"")
            .AppendLine()
            .AppendLine()
            .Append(converter.ConvertChildren(htmlNode))
            .AppendLine("::: zone-end")
            .ToString();
    }

    public override string Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        Debug.Assert(CanConvert(htmlInput));

        var classTags = htmlInput.GetAttributeValue("class", "").Split(' ');
        string? result = null;
        foreach (var classText in classTags)
        {
            if (_blockConverters.TryGetValue(classText, out var handler))
            {
                result = handler(classText, converter, htmlInput);
                break;
            }
        }

        if (result == nameof(IgnoreBlock))
            return string.Empty;
        if (string.IsNullOrWhiteSpace(result))
            result = converter.ConvertChildren(htmlInput);
        
        return result;
    }

    private static string QuoteBlock(string className, HtmlConverter converter, HtmlNode htmlNode)
    {
        var sb = new StringBuilder($"> [!{className.Trim().ToUpper()}]{Environment.NewLine}");
        foreach (var child in htmlNode.ChildNodes)
        {
            var lines = converter.Convert(child).Split('\n', '\r');
            bool blankLine = false;
            for (int i = 0; i < lines.Length; i++)
            {
                var text = lines[i].Trim();
                if (i == 0
                    && (string.IsNullOrWhiteSpace(text) || text.Equals(className, StringComparison.OrdinalIgnoreCase)))
                    continue;

                if (string.IsNullOrWhiteSpace(text))
                {
                    if (blankLine)
                        continue;
                    blankLine = true;
                }
                else
                    blankLine = false;
                
                sb.Append("> ")
                  .AppendLine(text);
            }
        }

        // Remove any ending blank lines.
        var result = sb.ToString();
        if (result.EndsWith($"> {Environment.NewLine}"))
        {
            result = result[..^4] + Environment.NewLine;
        }

        return result;
    }
}
