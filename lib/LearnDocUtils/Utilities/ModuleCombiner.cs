using System.Diagnostics;
using MSLearnRepos;

namespace LearnDocUtils;

public static class ModuleCombiner
{
    public static async Task<(Module module, string markdownFile)> DownloadModuleAsync(
        ILearnRepoService learnRepo, string learnFolder, string outputFolder, bool embedNotebooks)
    {
        if (string.IsNullOrEmpty(outputFolder))
            throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or empty.", nameof(outputFolder));

        // Load the module from the source.
        var module = await learnRepo.GetModuleAsync(learnFolder);
        if (module == null)
            throw new ArgumentException($"Failed to parse Learn module from {learnFolder}", nameof(learnFolder));

        // Load the units.
        await learnRepo.LoadUnitsAsync(module);

        // Make a local copy
        await learnRepo.CopyToAsync(module, outputFolder);

        // Combine all the units into a single Markdown file for us to process.
        var markdownFile = Path.Combine(outputFolder, Path.ChangeExtension(Path.GetFileNameWithoutExtension(learnFolder)??"units.g.",".md"));
        await using var tempFile = new StreamWriter(markdownFile);

        foreach (var unit in module.Units)
        {
            // Get the unit text
            string unitFn = unit.GetContentFilename();
            var mdText = !string.IsNullOrEmpty(unitFn)
                ? await File.ReadAllTextAsync(Path.Combine(outputFolder, unitFn))
                : unit.Content ?? "";

            // Pull out any YAML header (this is added for localization).
            string header = null;
            if (ExtractYamlHeader(ref mdText, out header))
            {
                // TODO: where should we put this?
            }
                
            // Write the title (H1)
            await tempFile.WriteLineAsync($"# {unit.Title}");
            await tempFile.WriteLineAsync();

            // Writ the content.
            if (!string.IsNullOrEmpty(mdText))
            {
                await tempFile.WriteLineAsync(mdText.Trim('\r', '\n'));
                await tempFile.WriteLineAsync();
            }

            // See if the content is a notebook and we are to embed them.
            if (!string.IsNullOrEmpty(unit.Notebook) && embedNotebooks)
            {
                string url = DetermineNotebookUrl(learnRepo, module.Url, unit.Notebook);
                var nbText = await NotebookConverter.Convert(url);
                if (nbText != null)
                {
                    await tempFile.WriteLineAsync(nbText.Trim('\r', '\n'));
                    await tempFile.WriteLineAsync();
                }
                else 
                    throw new Exception($"Failed to locate and download notebook {url}");
            }

            // Write any quiz
            if (unit.Quiz != null)
            {
                if (!string.IsNullOrEmpty(unit.Quiz.Title))
                    await tempFile.WriteLineAsync($"## {unit.Quiz.Title}{Environment.NewLine}");

                foreach (var question in unit.Quiz.Questions)
                {
                    await tempFile.WriteLineAsync($"### {question.Content}");
                    foreach (var choice in question.Choices)
                    {
                        await tempFile.WriteAsync(choice.IsCorrect ? "- [x] " : "- [ ] ");
                        await tempFile.WriteLineAsync(EscapeContent(choice.Content));
                        if (!string.IsNullOrEmpty(choice.Explanation))
                        {
                            await tempFile.WriteLineAsync($"    - {EscapeContent(choice.Explanation)}");
                        }
                    }
                    await tempFile.WriteLineAsync();
                }
            }
        }

        return (module, markdownFile);
    }

    static bool ExtractYamlHeader(ref string mdText, out string header)
    {
        const string YamlMarker = "---";
        header = null;

        using var reader = new StringReader(mdText);
        var line = reader.ReadLine();

        // Skip blanks. Shouldn't be any.
        while (line != null && string.IsNullOrWhiteSpace(line))
            line = reader.ReadLine();

        if (line == null) return false;

        if (line.Trim().StartsWith(YamlMarker))
        {
            line = reader.ReadLine();
            while (line != null && line.Trim().StartsWith(YamlMarker) == false)
            {
                header += line + Environment.NewLine;
                line = reader.ReadLine();
            }

            // Never found ending mark.
            if (line == null)
            {
                header = null;
                return false;
            }

            mdText = reader.ReadToEnd();
            return true;
        }

        return false;
    }

    private static string EscapeContent(string text) => text.Replace("{", @"\{");

    private static string DetermineNotebookUrl(ILearnRepoService learnRepo, string moduleUrl, string unitNotebook)
    {
        if (unitNotebook.ToLower().StartsWith("http"))
            return unitNotebook;

        // Look for an absolute link.
        if (unitNotebook.StartsWith('/'))
        {
            string path = "/learn/modules/";

            string[] moduleParts = moduleUrl.Split('/', '\\'); // support local and remote paths.
            path += moduleParts.Last(s => !string.IsNullOrEmpty(s));

            Debug.Assert(unitNotebook.StartsWith(path));
            unitNotebook = unitNotebook[(path.Length + 1)..];
        }

        if (!learnRepo.RootPath.ToLower().StartsWith("http")) 
            return Path.Combine(learnRepo.RootPath, unitNotebook);
            
        string url = moduleUrl;
        if (!moduleUrl.EndsWith('/'))
            url += '/';
        url += unitNotebook;
        return url;
    }
}