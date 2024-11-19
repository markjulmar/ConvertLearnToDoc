using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Metadata;

public static class DocsMetadataFactory
{
    private const string Conceptual = "Conceptual";
    
    public static async Task<IDocsMetadata> CreateAsync(HtmlDocument htmlDoc)
    {
        var pageType = htmlDoc.DocumentNode.SelectSingleNode("//head/meta[@name='page_type']").GetAttributeValue("content", "");
        if (pageType == "conceptual")
            return DocsMetadata.Create(htmlDoc);
        if (pageType == "learn")
        {
            var pageKind = htmlDoc.DocumentNode.SelectSingleNode("//head/meta[@name='page_kind']").GetAttributeValue("content", "");
            if (pageKind == "unit")
                return await UnitMetadata.Create(htmlDoc);
            if (pageKind == "module")
                return await ModuleMetadata.Create(htmlDoc);
        }

        throw new ArgumentException("Unable to determine the document type from the URL.", nameof(htmlDoc));
    }
}