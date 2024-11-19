using System.Text;
using HtmlAgilityPack;
using Julmar.DocsToMarkdown.Converters;
using Julmar.DocsToMarkdown.Metadata;

namespace Julmar.DocsToMarkdown;

internal sealed class HtmlConverter(HtmlDocument htmlDocument) : IHtmlConverter
{
    private IDocsMetadata? _metadata;
    readonly List<IConverter> _converters = [
            new Header(), new Anchor(), new Image(), new Paragraph(),
            new List(), new Table(), new Code(), new Form(),
            new Bold(), new Emphasis(),
            new DivSpan()
        ];

    public Action<string>? UnhandledTagCallback { get; set; } = null;
    
    public bool UseDocfxExtensions { get; set; }
    public IDocsMetadata Metadata => _metadata!;
    public IConverter? GetConverterFor(HtmlNode node)
    {
        return _converters.FirstOrDefault(c => c.CanConvert(node));
    }

    public IEnumerable<string> Process()
    {
        return _metadata!.PageKind switch
        {
            // Article?
            "conceptual" or "unit" => ProcessOneDocument(htmlDocument),
            "module" => ProcessTrainingModule(htmlDocument),
            _ => throw new NotSupportedException()
        };
    }

    private static IEnumerable<string> ProcessTrainingModule(HtmlDocument document)
    {
        var url = document.DocumentNode.SelectSingleNode("//head/meta[@property='og:url']").GetAttributeValue("content", "");
        if (string.IsNullOrEmpty(url))
            throw new ArgumentException("Cannot find og:url meta tag to parse.");
        
        var unitList = document.GetElementbyId("unit-list");
        if (unitList == null)
            throw new ArgumentException("Cannot find unit-list div to parse.");
        var units = unitList.Descendants("a").ToList();
        foreach (var anchor in units)
        {
            var href = anchor.GetAttributeValue("href", "");
            if (!string.IsNullOrWhiteSpace(href))
            {
                var uri = !href.StartsWith("http")
                    ? new Uri(new Uri(url), href)
                    : new Uri(href);
                var unitDoc = HtmlConverterFactory.Create(uri).Result;
                foreach (var line in unitDoc.Process())
                {
                    yield return line;
                }

                yield return Environment.NewLine;
            }
        }
    }

    private IEnumerable<string> ProcessOneDocument(HtmlDocument document)
    {
        var root = document.DocumentNode.SelectSingleNode("//div[@class='content']") 
                   ?? document.DocumentNode.SelectSingleNode("//div[@class='content ']");
        if (root == null)
            throw new ArgumentException("Cannot find content div to parse.");
        
        foreach (var node in root.ChildNodes)
        {
            var result = Convert(node);
            if (!string.IsNullOrWhiteSpace(result))
            {
                yield return result;
            }
            else if (node.NodeType == HtmlNodeType.Element)
            {
                UnhandledTagCallback?.Invoke(node.OuterHtml);
            }
        }
    }
    
    public string Convert(HtmlNode node) 
        => GetConverterFor(node)?.Convert(this, node) ?? "";

    private static IConverter[] InlineConverters => [new Text(), new Bold(), new Emphasis(), new Anchor()];
    
    internal string ConvertChildren(HtmlNode htmlInput, params IConverter[] additionalConverters)
    {
        var sb = new StringBuilder();
        foreach (var child in htmlInput.ChildNodes)
        {
            bool foundConverter = false;
            var result = Convert(child);
            if (!string.IsNullOrWhiteSpace(result))
            {
                foundConverter = true;
                sb.Append(result);
                if (IsBlockElement(child))
                    sb.AppendLine();
            }
            else
            {
                foreach (var cvt in additionalConverters.Union(InlineConverters))
                {
                    if (cvt.CanConvert(child))
                    {
                        result = cvt.Convert(this, child);
                        if (!string.IsNullOrWhiteSpace(result))
                        {
                            foundConverter = true;
                            sb.Append(result);
                            break;
                        }
                    }
                }
            }

            if (!foundConverter)
            {
                if (IsBlockElement(child))
                    UnhandledTagCallback?.Invoke(child.OuterHtml);
            }
        }

        return sb.ToString();
    }

    private static readonly string[] BlockElements = ["div", "p", "ul", "ol", "table", "h1", "h2", "h3", "h4", "h5", "h6"]; 
    private static bool IsBlockElement(HtmlNode child)
    {
        var tagName = child.Name.ToLowerInvariant();
        return BlockElements.Contains(tagName);
    }

    internal async Task<IHtmlConverter> Initialize()
    {
        _metadata = await DocsMetadataFactory.CreateAsync(htmlDocument);
        return this;
    }
}