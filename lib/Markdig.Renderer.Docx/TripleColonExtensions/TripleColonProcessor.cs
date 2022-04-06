using System.Drawing;
using IContainer = DXPlus.IContainer;

namespace Markdig.Renderer.Docx.TripleColonExtensions;

internal static class TripleColonProcessor
{
    public static void Write(IDocxObjectRenderer renderer, MarkdownObject obj,
        IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
    {
        switch (extension.Extension.Name)
        {
            case "image":
                HandleImage(owner, document, currentParagraph, extension);
                break;
            case "zone":
                HandleZonePivot(renderer, (ContainerBlock)obj, owner, document, currentParagraph, extension);
                break;
            case "code":
                HandleCode(renderer, (ContainerBlock) obj, owner, document, currentParagraph, extension);
                break;
            default:
                break;
        }
    }

    private static void HandleCode(IDocxObjectRenderer renderer, ContainerBlock containerBlock, IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
    {
        var language = extension.Attributes["language"];
        var source = extension.Attributes["source"];
        extension.Attributes.TryGetValue("range", out var range);
        extension.Attributes.TryGetValue("highlight", out var highlight);

        var p = currentParagraph ?? document.AddParagraph();
        p.AddText(new Run($"{{codeBlock: language={language}, source=\"{source}\", range={range}, highlight={highlight}}}",
            new Formatting {Highlight = Highlight.Blue, Color = Color.White}));
        if (currentParagraph == null) p.Newline();
    }

    private static void HandleZonePivot(IDocxObjectRenderer renderer, ContainerBlock block, 
        IDocxRenderer owner, IDocument document, Paragraph currentParagraph, TripleColonElement extension)
    {
        var pivot = extension.Attributes["pivot"];
        if (owner.ZonePivot == null
            || pivot != null && pivot.ToLower().Contains(owner.ZonePivot))
        {
            if (owner.ZonePivot == null)
            {
                var p = currentParagraph ?? document.AddParagraph();
                p.AddText(new Run($"{{zonePivot: \"{pivot}\"}}", new Formatting {Highlight = Highlight.Red, Color = Color.White }));
                if (currentParagraph == null) p.Newline();
            }
                
            renderer.WriteChildren(block, owner, document, currentParagraph);
                
            if (owner.ZonePivot == null)
            {
                var p = currentParagraph ?? document.AddParagraph();
                p.AddText(new Run($"{{end-zonePivot: \"{pivot}\"}}",
                    new Formatting {Highlight = Highlight.Red, Color = Color.White}));
                if (currentParagraph == null) p.Newline();
            }
        }
    }

    private static void HandleImage(IDocxRenderer owner, IContainer document, Paragraph currentParagraph, TripleColonElement extension)
    {
        currentParagraph ??= document.AddParagraph();

        extension.Attributes.TryGetValue("type", out var type);
        extension.Attributes.TryGetValue("alt-text", out var title);
        extension.Attributes.TryGetValue("loc-scope", out var localization);
        extension.Attributes.TryGetValue("source", out var source);
        extension.Attributes.TryGetValue("border", out var hasBorder);
        extension.Attributes.TryGetValue("lightbox", out var lightboxImageUrl);
        extension.Attributes.TryGetValue("link", out var link);

        string description = null;
        if (extension.Container?.Count > 0 && type == "complex")
        {
            // Should be strictly text as this is for screen readers.
            description = string.Join("\r\n", extension.Container.Select(b => (b as ParagraphBlock)?.Inline)
                .SelectMany(ic => ic.Select(il => il.ToString())));
        }

        var drawing = owner.InsertImage(currentParagraph, extension.Container ?? (MarkdownObject)extension.Inlines, source, title, description, hasBorder?.ToLower()=="true");
        if (drawing != null)
        {
            string commentText = Globals.UseExtension;
            if (!string.IsNullOrEmpty(link))
                commentText += $" link:\"{link}\"";
            if (!string.IsNullOrEmpty(localization))
                commentText += $" loc-scope:{localization}";
            if (!string.IsNullOrEmpty(lightboxImageUrl))
                commentText += $" lightbox:\"{lightboxImageUrl}\"";

            owner.AddComment(currentParagraph, commentText);
        }
    }
}