namespace Docx.Renderer.Markdown.Renderers;

internal static class RenderHelpers
{
    internal static bool CollapseEmptyTags(Paragraph block)
    {
        // Fix cases where Word traps an ending space in the emphasis and then
        // closes the emphasis before more text. This is illegal in Markdown
        // so we'll fudge it a little and move the space to the next block.

        string capturedSpaces = null;

        var lastBlock = block.LastOrDefault();
        if (lastBlock is BoldText or ItalicText or BoldItalicText)
        {
            var dt = (Text)lastBlock;

            // Just whitespace?
            if (string.IsNullOrWhiteSpace(dt.Text))
            {
                // Kill the block completely.
                capturedSpaces = dt.Text;
                block.Remove(dt);
            }
            else if (dt.Text.EndsWith(' '))
            {
                dt.Text = dt.Text[..^1];
                capturedSpaces = " ";
            }
        }

        if (capturedSpaces != null)
        {
            block.Add(new Text(capturedSpaces));
            return true;
        }

        return false;
    }
}