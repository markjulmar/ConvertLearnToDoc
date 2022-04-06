namespace Docx.Renderer.Markdown.Renderers;

public abstract class MarkdownObjectRenderer<T> : IMarkdownObjectRenderer
{
    public bool CanRender(object element) => element is T;

    void IMarkdownObjectRenderer.Render(IMarkdownRenderer renderer,
        MarkdownDocument document, MarkdownBlock blockOwner,
        object elementToRender, RenderBag tags) =>
        Render(renderer, document, blockOwner, (T) elementToRender, tags);
        
    protected abstract void Render(IMarkdownRenderer renderer,
        MarkdownDocument document, MarkdownBlock blockOwner,
        T elementToRender, RenderBag tags);

    /// <summary>
    /// Look for a specific comment on the given paragraph. We assume the keys are space delimited
    /// and quoted if they have spaces. Surrounding quotes are removed. If the key has no value, then the
    /// key itself is returned. If the key doesn't exist, null is returned.
    /// </summary>
    /// <param name="paragraph"></param>
    /// <param name="prefix"></param>
    /// <returns></returns>
    protected static string FindCommentValue(DXParagraph paragraph, string prefix)
    {
        if (paragraph != null)
        {
            var comments = paragraph.Comments.SelectMany(c => c.Comment.Paragraphs.Select(p => p.Text ?? ""));
            var found = comments.FirstOrDefault(c => c.Contains(prefix, StringComparison.InvariantCultureIgnoreCase));
            if (!string.IsNullOrEmpty(found))
            {
                found = found.TrimEnd('\r', '\n');

                int start = found.IndexOf(prefix, StringComparison.InvariantCultureIgnoreCase) + prefix.Length;
                if (start == found.Length || found[start] == ' ')
                    return prefix;

                Debug.Assert(found[start] == ':');
                start++;
                char lookFor = ' ';
                if (found[start] == '\"')
                {
                    lookFor = '\"';
                    start++;
                }

                int end = start;
                while (end < found.Length)
                {
                    if (found[end] == lookFor)
                        break;
                    end++;
                }

                Debug.Assert(start >= 0 && found.Length > start);
                return found[start..end];
            }
        }

        return null;
    }

}