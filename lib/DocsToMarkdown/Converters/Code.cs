using System.Text;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal class Code(): BaseConverter("code", "pre")
{
    public override string? Convert(HtmlConverter converter, HtmlNode htmlInput)
    {
        if (htmlInput.Name == "pre")
        {
            var codeBlock = htmlInput.SelectSingleNode("code");
            if (codeBlock == null)
            {
                return new StringBuilder("<pre>")
                    .AppendLine()
                    .Append(converter.ConvertChildren(htmlInput))
                    .AppendLine("</pre>")
                    .AppendLine()
                    .ToString();
            }

            htmlInput = codeBlock;
        }
        return ProcessCodeBlock(converter, htmlInput);
    }

    private static string ProcessCodeBlock(HtmlConverter converter, HtmlNode codeBlock)
    {
        var className = codeBlock.GetAttributeValue("class", "").Trim();
        if (className.StartsWith("lang-"))
            className = className[5..];

        var code = System.Net.WebUtility.HtmlDecode(codeBlock.InnerText);
        if (code.Contains('\n'))
        {
            return new StringBuilder("```")
                .Append(className)
                .AppendLine()
                .AppendLine(code.TrimEnd('\r', '\n'))
                .AppendLine("```")
                .AppendLine()
                .ToString();
        }

        return $"`{code}`";
    }
}