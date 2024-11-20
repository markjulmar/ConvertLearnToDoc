using System.Text.RegularExpressions;

namespace LearnDocUtils;

internal static partial class DocToMarkdownRenderer
{
    /// <summary>
    /// This is the list of unicode space characters we see
    /// in Word docs. This includes non-breaking spaces, narrow spaces,
    /// and zero width spaces.
    /// </summary>
    static readonly (string find,string replace)[] Replacements =
    {
        ("\xa0", " "), 
        ("\u200B", ""), 
        ("\u202F", " "),
        ("{tabgroup-end}", "---")
    };

    /// <summary>
    /// Do some post-conversion cleanup of markers, paths, and triple-colon placeholders.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    internal static string PostProcessMarkdown(string text)
    {
        text = text.Trim('\r', '\n').Trim('\n', '\r');
        text = Replacements.Aggregate(text, (current, ch) => current.Replace(ch.find, ch.replace));

        // Do all the doc replacements we did from an original doc.
        text = RegionRegex().Replace(text, m => $"<rgn>{m.Groups[1].Value.Trim()}</rgn>");
        text = BeginZonePivotRegex().Replace(text, m => $":::zone pivot={m.Groups[1].Value.Trim()}");
        text = EndZonePivotRegex().Replace(text, m => $":::zone-end");
        text = IncludeRegex().Replace(text, m => $"[!include[]({m.Groups[1].Value.Trim()})]");
        text = TabGroupRegex().Replace(text, m => $"# [{m.Groups[1].Value.Trim()}](#tab/{m.Groups[2].Value.Trim()})]");

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

    [GeneratedRegex(@"{rgn (.*?)}")]
    private static partial Regex RegionRegex();

    [GeneratedRegex(@"{zonePivot:(.*?)}")]
    private static partial Regex BeginZonePivotRegex();

    [GeneratedRegex(@"{end-zonePivot:(.*?)}")]
    private static partial Regex EndZonePivotRegex();
    
    [GeneratedRegex(@"{include ""(.*?)"".*}")]
    private static partial Regex IncludeRegex();
    
    [GeneratedRegex(@"{tabgroup ""(.*?)"" (.*?)}")]
    private static partial Regex TabGroupRegex();
}