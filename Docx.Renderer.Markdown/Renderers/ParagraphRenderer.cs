using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Julmar.GenMarkdown;
using DXParagraph = DXPlus.Paragraph;
using Header = Julmar.GenMarkdown.Header;
using Paragraph = Julmar.GenMarkdown.Paragraph;

namespace Docx.Renderer.Markdown.Renderers
{
    public sealed class ParagraphRenderer : MarkdownObjectRenderer<DXParagraph>
    {
        private static readonly Dictionary<string, 
            Action<IMarkdownRenderer,MarkdownDocument,MarkdownBlock,DXParagraph, RenderBag>> creators = new()
        {
            { "Heading1", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { "Heading2", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { "Heading3", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { "Heading4", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { "Heading5", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { @"(?s:.*)Code", CreateCodeBlock },
            { @"(?s:.*)Quote", CreateBlockQuote },
            { ".*", CreateParagraph } // Must be last.
        };

        private static readonly Dictionary<string, string> blockHeaders = new Dictionary<string, string>
        {
            {"note:", "NOTE"},
            {"tip:", "TIP"},
            {"warning", "WARNING"},
            {"error", "ERROR"}
        };
        
        protected override void Render(IMarkdownRenderer renderer, 
            MarkdownDocument document, MarkdownBlock blockOwner, 
            DXParagraph element, RenderBag tags)
        {
            tags ??= new RenderBag(); 
            
            var tf = tags.Get<TextFormatting>(nameof(TextFormatting));
            tf.StyleName = element.Properties.StyleName;
            if (element.Properties.DefaultFormatting.Bold)
                tf.Bold = true;
            if (element.Properties.DefaultFormatting.Italic)
                tf.Italic = true;
            if (TextFormatting.IsMonospaceFont(element.Properties.DefaultFormatting.Font))
                tf.Monospace = true;
            
            tags.AddOrReplace(nameof(TextFormatting), tf);

            if (tf.Monospace)
                CreateCodeBlock(renderer, document, blockOwner, element, tags);
            else
            {
                if (!creators.TryGetValue(tf.StyleName, out var creator))
                    creator = creators.FirstOrDefault(kvp => 
                        Regex.IsMatch(tf.StyleName, kvp.Key, RegexOptions.IgnoreCase)).Value;
                Debug.Assert(creator != null);
                creator.Invoke(renderer, document, blockOwner, element, tags);
            }
        }

        private static void CreateHeader(int level, IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            Debug.Assert(blockOwner == null);
            
            var p = new Header(level);
            document.Add(p);
            renderer.WriteContainer(document, p, element.Runs, tags);
        }

        private static void CreateBlockQuote(IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            Debug.Assert(blockOwner == null);

            bool newBlockQuote = false;
            var blockQuote = document.LastOrDefault() as BlockQuote;
            if (blockQuote == null)
            {
                blockQuote = new BlockQuote();
                document.Add(blockQuote);
                newBlockQuote = true;
            }
            else
            {
                if (blockQuote.LastOrDefault() is Paragraph { Count: > 0 } p)
                {
                    blockQuote.Add(new Paragraph());
                }
            }

            var runs = element.Runs.ToList();
            for (int i = 0; i < runs.Count; i++)
            {
                bool processed = false;
                var run = runs[i];
                if (newBlockQuote)
                {
                    foreach (var kvp in 
                        from kvp in blockHeaders 
                        let text = run.Text 
                        where run.Text.Contains(kvp.Key, StringComparison.CurrentCultureIgnoreCase) 
                        select kvp)
                    {
                        blockQuote.Add($"[!{kvp.Value}]");
                        blockQuote.Add(new Paragraph());
                        processed = true;

                        // Skip any blanks after that.
                        while (i+1 < runs.Count)
                        {
                            if (runs[i + 1].HasText
                                && string.IsNullOrWhiteSpace(runs[i + 1].Text))
                                i++;
                            else break;
                        }
                        break;
                    }
                }

                if (!processed)
                {
                    var p = blockQuote.LastOrDefault();
                    if (p == null)
                    {
                        p = new Paragraph();
                        blockQuote.Add(p);
                    }
                    
                    renderer.WriteElement(document, p, run, tags);
                }
            }
        }

        private static void CreateParagraph(IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            if (blockOwner == null)
            {
                blockOwner = new Paragraph();
                document.Add(blockOwner);
            }
            renderer.WriteContainer(document, blockOwner, element.Runs, tags);
        }

        private static void CreateCodeBlock(IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            Debug.Assert(blockOwner == null);

            document.Add(new CodeBlock(element.Text));
        }
    }
}
