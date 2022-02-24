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

        for (int index = 0; index < result.OldText.Lines.Count; index++)
        {
            var originalDp = result.OldText.Lines[index];
            var newDp = result.NewText.Lines[index];
            if (originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Unchanged
                && newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Unchanged)
                continue;

            var odp = originalDp.SubPieces
                .Where(odp =>
                    odp.Type != DiffPlex.DiffBuilder.Model.ChangeType.Unchanged &&
                    odp.Type != DiffPlex.DiffBuilder.Model.ChangeType.Imaginary)
                .ToList();
            var ndp = newDp.SubPieces
                .Where(ndp =>
                    ndp.Type != DiffPlex.DiffBuilder.Model.ChangeType.Unchanged &&
                    ndp.Type != DiffPlex.DiffBuilder.Model.ChangeType.Imaginary)
                .ToList();

            const string listMarkers = "-*", emphasisChars = "*_";
            if (odp.Count > 0 && odp.Count == ndp.Count)
            {
                bool skip = true;
                for (int i = 0; i < odp.Count; i++)
                {
                    if (odp[i].Position != ndp[i].Position)
                    {
                        skip = false;
                        break;
                    }

                    string original = odp[i].Text;
                    string newt = ndp[i].Text;

                    if (original.Length != newt.Length)
                    {
                        skip = false;
                        break;
                    }

                    // Ignore Markdown differences involving the emphasis or bullet.
                    if (Enumerable.Range(0, original.Length).All(
                            n => original[n] == newt[n]
                                 || (listMarkers.Contains(original[n]) && listMarkers.Contains(newt[n]))
                                 || (emphasisChars.Contains(original[n]) && emphasisChars.Contains(newt[n]))))
                        continue;

                    skip = false;
                    break;
                }

                if (skip) continue;
            }

            // Skip blank lines inserted or deleted. We assume this doesn't affect content.
            if ((originalDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Deleted && originalDp.Text.Length == 0)
                || (newDp.Type == DiffPlex.DiffBuilder.Model.ChangeType.Inserted && newDp.Text.Length == 0))
                continue;

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
