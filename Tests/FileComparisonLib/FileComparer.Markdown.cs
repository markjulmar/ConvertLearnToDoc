using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;

namespace FileComparisonLib;

public static partial class FileComparer
{
    public static IEnumerable<Difference> Markdown(string fn1, string fn2)
    {
        var originalText = File.ReadAllText(fn1);
        var newText = File.ReadAllText(fn2);

        var differ = new Differ();
        var sbsDiffBuilder = new SideBySideDiffBuilder(differ);
        var result = sbsDiffBuilder.BuildDiffModel(originalText, newText, true);

        List<int> processed = new List<int>();

        for (int index = 0; index < result.OldText.Lines.Count; index++)
        {
            var originalDp = result.OldText.Lines[index];
            var newDp = result.NewText.Lines[index];

            // Unchanged line?
            if (originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Unchanged
                && newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Unchanged)
                continue;

            // Skip blank lines inserted or deleted. We assume this doesn't affect content.
            if (originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Deleted && originalDp.Text.Length == 0
                || newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Inserted && newDp.Text.Length == 0)
                continue;

            // If the text was deleted, see if we can find an insertion in the new text to match it to.
            if (originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Deleted
                && newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Imaginary)
            {
                // Line was already processed as a deletion.
                if (processed.Contains(-1 * index))
                    continue;

                var addCheck = result.NewText.Lines.FirstOrDefault(l =>
                    l.Position == originalDp.Position && l.Type == DiffPlex.DiffBuilder.Model.ChangeType.Inserted);
                if (addCheck != null)
                {
                    newDp = addCheck;
                    processed.Add(result.NewText.Lines.IndexOf(addCheck));
                }
            }
            // Same for the other side - if we have an addition, then see if we can match it to a delete
            // on the other file.
            else if (originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Imaginary
                     && newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Inserted)
            {
                // Line was already processed as an addition.
                if (processed.Contains(index))
                    continue;

                var delCheck = result.OldText.Lines.FirstOrDefault(l =>
                    l.Position == newDp.Position && l.Type == DiffPlex.DiffBuilder.Model.ChangeType.Deleted);
                if (delCheck != null)
                {
                    originalDp = delCheck;
                    processed.Add(result.OldText.Lines.IndexOf(delCheck) * -1);
                }
            }

            // Ignore Markdown differences involving the emphasis or bullet.
            // These changes don't impact the rendering and can be interchanged.
            const string listMarkers = "-*", emphasisChars = "*_";
            string original = originalDp.Text, newt = newDp.Text;
            if (original != null && original.Length == newt?.Length
                && Enumerable.Range(0, original.Length).All(
                    n => original[n] == newt[n]
                         || listMarkers.Contains(original[n]) && listMarkers.Contains(newt[n])
                         || emphasisChars.Contains(original[n]) && emphasisChars.Contains(newt[n])))
                continue;

            // Found a diff.
            yield return CreateDiff(originalDp, newDp);
        }
    }

    private static MarkdownDiff CreateDiff(DiffPiece original, DiffPiece changed)
    {
        var diff = new MarkdownDiff
        {
            OriginalPosition = original.Position,
            NewPosition = changed.Position,
            OriginalValue = original.Text,
            NewValue = changed.Text
        };

        if (original.Type == DiffPlex.DiffBuilder.Model.ChangeType.Imaginary)
        {
            diff.Change = changed.Type switch
            {
                DiffPlex.DiffBuilder.Model.ChangeType.Deleted => ChangeType.Deleted,
                DiffPlex.DiffBuilder.Model.ChangeType.Inserted => ChangeType.Added,
                _ => ChangeType.Changed
            };
        }
        else
        {
            diff.Change = original.Type switch
            {
                DiffPlex.DiffBuilder.Model.ChangeType.Deleted => ChangeType.Deleted,
                DiffPlex.DiffBuilder.Model.ChangeType.Inserted => ChangeType.Added,
                _ => ChangeType.Changed
            };
        }

        return diff;
    }
}
