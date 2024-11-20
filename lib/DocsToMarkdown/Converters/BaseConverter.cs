using System.Collections.ObjectModel;
using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Converters;

internal abstract class BaseConverter : IConverter
{
    protected readonly ReadOnlyCollection<string> SupportedTags; 
    
    protected BaseConverter(string tagName, params string[] alternatives)
    {
        if (alternatives.Length > 0)
        {
            var tagNames = new string[alternatives.Length + 1];
            tagNames[0] = tagName;
            Array.Copy(alternatives, 0, tagNames, 1, alternatives.Length);
            SupportedTags = Array.AsReadOnly(tagNames);
        }
        else
        {
            SupportedTags = Array.AsReadOnly([tagName]);
        }
    }
    
    protected BaseConverter()
    {
        SupportedTags = Array.AsReadOnly(Array.Empty<string>());
    }

    public virtual bool CanConvert(HtmlNode htmlInput) 
        => SupportedTags.IndexOf(htmlInput.Name.ToLower()) >= 0;

    public abstract string? Convert(HtmlConverter converter, HtmlNode htmlInput);

    protected string SimplifyRelativePath(IHtmlConverter converter, string source)
    {
        if (source.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            return source;

        // Try to massage source.
        source = source.Replace('\\', '/').ToLower();
        if (source.StartsWith("./"))
            source = source[2..];
        else if (source.StartsWith('/'))
            source = source[1..];
        else if (source.StartsWith("../"))
        {
            // Special case
            var moduleName =
                converter.IsTrainingModulePage
                    ? converter.Url.Segments.Last().TrimEnd('/')
                    : converter.Metadata.PageKind == "unit"
                        ? converter.Url.Segments[^2].TrimEnd('/')
                        : null;
            var sections = source.ToLower().Split('/');
            int start = moduleName != null
                ? Array.IndexOf(sections, moduleName) + 1
                : 0;
            while (sections[start] == "..")
                start++;
            source = string.Join('/', sections, start, sections.Length - start);
        }

        return source;
    }
}