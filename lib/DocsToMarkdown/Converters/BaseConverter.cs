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
}