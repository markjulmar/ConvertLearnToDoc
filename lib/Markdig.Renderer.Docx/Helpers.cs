namespace Markdig.Renderer.Docx;

internal static class Helpers
{
    // Some of the Learn content has invalid characters such as '\v' in the pages. This works in the Markdown
    // but fails the strict parsing required by XML and docx.
    public static string CleanText(string rawText) => new(rawText.Where(XmlConvert.IsXmlChar).ToArray());

    /// <summary>
    /// Returns the text of all literals following the given inline element until a non-literal is found.
    /// </summary>
    /// <param name="renderer"></param>
    /// <param name="item"></param>
    /// <returns></returns>
    public static string ReadLiteralTextAfterTag(IDocxRenderer renderer, Inline item)
    {
        StringBuilder sb = new();
        if (!item.IsClosed)
        {
            while (item.NextSibling is LiteralInline li)
            {
                sb.Append(CleanText(li.Content.ToString()));
                renderer.OutOfPlaceRendered.Add(li);
                item = item.NextSibling;
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Returns the HTML tag associated with a string of HTML.
    /// </summary>
    /// <param name="htmlTag"></param>
    /// <returns></returns>
    public static string GetTag(string htmlTag)
    {
        if (string.IsNullOrEmpty(htmlTag))
            return null;

        int startPos = 1;
        if (htmlTag.StartsWith("</"))
            startPos = 2;
        int endPos = startPos;
        while (char.IsLetter(htmlTag[endPos]))
            endPos++;
        return htmlTag.Substring(startPos, endPos - startPos).ToLower();
    }

}