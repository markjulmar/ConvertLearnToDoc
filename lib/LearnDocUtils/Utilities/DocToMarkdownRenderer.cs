using System.Text.RegularExpressions;

namespace LearnDocUtils;

internal static class DocToMarkdownRenderer
{
    /// <summary>
    /// Do some post-conversion cleanup of markers, paths, and triple-colon placeholders.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    internal static string PostProcessMarkdown(string text)
    {
        text = text.Trim('\r').Trim('\n');
        text = text.Replace("\xa0", " "); // TODO: should this be &nbsp;?
        text = text.Replace("{tabgroup-end}", "---");

        text = Regex.Replace(text, @"{rgn (.*?)}", m => $"<rgn>{m.Groups[1].Value.Trim()}</rgn>");
        text = Regex.Replace(text, @"{zonePivot:(.*?)}", m => $":::zone pivot={m.Groups[1].Value.Trim()}");
        text = Regex.Replace(text, @"{end-zonePivot:(.*?)}", m => $":::zone-end");
        text = Regex.Replace(text, @"{include ""(.*?)"".*}", m => $"[!include[]({m.Groups[1].Value.Trim()})]");
        text = Regex.Replace(text, @"{tabgroup ""(.*?)"" (.*?)}", m => $"# [{m.Groups[1].Value.Trim()}](#tab/{m.Groups[2].Value.Trim()})]");

        // WWL templates use prefixes, convert these to our quote block note types.
        text = ProcessNotes(text, "note: ", "NOTE");
        text = ProcessNotes(text, "(*) ", "TIP");

        return text;
    }

    private static string ProcessNotes(string text, string lookFor, string header)
    {
        int index = text.IndexOf(lookFor, StringComparison.CurrentCultureIgnoreCase);
        while (index > 0)
        {
            if (text[index - 1] == '\r' || text[index - 1] == '\n')
            {
                int end = text.IndexOf('\n', index);
                if (end > index)
                {
                    end++;
                    string noteBlock = text.Substring(index + lookFor.Length, end - index - lookFor.Length).TrimEnd('\r', '\n');
                    text = text.Remove(index, end - index)
                        .Insert(index, $"> [!{header.ToUpper()}]{Environment.NewLine}> {noteBlock}{Environment.NewLine}");
                }
            }

            index = text.IndexOf(lookFor, index + 1, StringComparison.CurrentCultureIgnoreCase);
        }

        return text;
    }
}