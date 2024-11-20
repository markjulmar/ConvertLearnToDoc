using HtmlAgilityPack;

namespace Julmar.DocsToMarkdown.Metadata;

public static class DocsMetadataFactory
{
    public static async Task<IDocsMetadata> CreateAsync(HtmlDocument htmlDoc)
    {
        var pageType = htmlDoc.DocumentNode.SelectSingleNode("//head/meta[@name='page_type']").GetAttributeValue("content", "");
        switch (pageType)
        {
            case "conceptual":
                return new DocsMetadata(htmlDoc);
            case "learn":
            {
                var pageKind = htmlDoc.DocumentNode.SelectSingleNode("//head/meta[@name='page_kind']").GetAttributeValue("content", "");
                switch (pageKind)
                {
                    case "unit":
                        return await UnitMetadata.Create(htmlDoc);
                    case "module":
                        return await ModuleMetadata.Create(htmlDoc);
                }
                break;
            }
        }

        throw new ArgumentException("Unable to determine the document type from the URL.", nameof(htmlDoc));
    }
}