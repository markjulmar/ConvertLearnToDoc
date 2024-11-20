using System.Text;
using HtmlAgilityPack;
using Julmar.DocsToMarkdown.Converters;
using Julmar.DocsToMarkdown.Metadata;

namespace Julmar.DocsToMarkdown;

internal sealed class HtmlConverter(Uri url, HtmlDocument htmlDocument, HtmlConverterOptions? options) : IHtmlConverter
{
    private IDocsMetadata? _metadata;
    readonly List<IConverter> _converters = [
            new Header(), new Anchor(), new Image(), new Paragraph(),
            new List(), new Table(), new Code(), new Form(),
            new Bold(), new Emphasis(),
            new DivSpan()
        ];

    public IDocsMetadata Metadata => _metadata!;

    public bool IsTrainingModulePage => _metadata!.PageKind == "module";

    private IConverter? GetConverterFor(HtmlNode node) 
        => _converters.FirstOrDefault(c => c.CanConvert(node));

    public Uri Url => url;

    public async IAsyncEnumerable<IHtmlConverter> GetTrainingUnitsConverter()
    {
        if (!IsTrainingModulePage)
        {
            yield return this;
        }
        else
        {
            var unitList = htmlDocument.GetElementbyId("unit-list");
            if (unitList == null)
                throw new ArgumentException("Cannot find unit-list div to parse.");
            var units = unitList.Descendants("a").ToList();
            foreach (var uri in units
                         .Select(unit => unit.GetAttributeValue("href", ""))
                         .Where(href => !string.IsNullOrWhiteSpace(href))
                         .Select(anchor => !anchor.StartsWith("http")
                             ? new Uri(Url.AbsoluteUri.EndsWith('/') 
                                        ? Url : new Uri(Url.AbsoluteUri + "/"), anchor)
                             : new Uri(anchor)))
            {
                yield return await IHtmlConverter.LoadAsync(uri, options);
            }
        }
    }
    
    public IEnumerable<string> GetMarkdown()
    {
        // No data for modules.
        if (Metadata.PageKind == "module")
            yield break;
        
        var root = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='content']") 
                   ?? htmlDocument.DocumentNode.SelectSingleNode("//div[@class='content ']");
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
                options?.UnhandledTagCallback?.Invoke(node.OuterHtml);
            }
        }
    }

    internal bool UseDocfxExtensions => options?.UseDocfxExtensions ?? false;
    internal string Convert(HtmlNode node) => GetConverterFor(node)?.Convert(this, node) ?? "";
    internal void DownloadAsset(string source, Uri uri) => options?.DownloadAsset?.Invoke(source, uri);
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
                    options?.UnhandledTagCallback?.Invoke(child.OuterHtml);
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