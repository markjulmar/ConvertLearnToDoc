﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using DXPlus;
using Julmar.GenMarkdown;
using DXParagraph = DXPlus.Paragraph;
using Paragraph = Julmar.GenMarkdown.Paragraph;

namespace Docx.Renderer.Markdown.Renderers
{
    public sealed class ParagraphRenderer : MarkdownObjectRenderer<DXParagraph>
    {
        private static readonly Dictionary<string, 
            Action<IMarkdownRenderer,MarkdownDocument,MarkdownBlock,DXParagraph, RenderBag>> creators = new()
        {
            { "Heading1", (r,d,bo,e,tags) => CreateHeader(1,r,d,bo,e,tags) },
            { "Heading2", (r,d,bo,e,tags) => CreateHeader(2,r,d,bo,e,tags) },
            { "Heading3", (r,d,bo,e,tags) => CreateHeader(3,r,d,bo,e,tags) },
            { "Heading4", (r,d,bo,e,tags) => CreateHeader(4,r,d,bo,e,tags) },
            { "Heading5", (r,d,bo,e,tags) => CreateHeader(5,r,d,bo,e,tags) },
            { "ListParagraph", CreateListBlock },
            { @"(?s:.*)Code", CreateCodeBlock },
            { @"(?s:.*)Quote", CreateBlockQuote },
            { "CodeFooter", (_, _, _, _, _) => { /* Do nothing */ } },
            { ".*", CreateParagraph } // Must be last.
        };

        /// <summary>
        /// The quote block headers
        /// </summary>
        private static readonly Dictionary<string, string> blockHeaders = new(StringComparer.CurrentCultureIgnoreCase)
        {
            {"note", "NOTE"},
            {"tip", "TIP"},
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
            if (element.Properties.DefaultFormatting.CapsStyle == CapsStyle.SmallCaps)
                tf.KbdTag = true;
            
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

            var p = new Heading(level);
            document.Add(p);
            renderer.WriteContainer(document, p, element.Runs, tags);
        }

        private static void CreateBlockQuote(IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            Debug.Assert(blockOwner == null);

            bool newBlockQuote = false;
            if (document.LastOrDefault() is BlockQuote blockQuote)
            {
                if (blockQuote.LastOrDefault() is Paragraph {Count: > 0} p)
                {
                    blockQuote.Add(new Paragraph());
                }
            }
            else
            {
                blockQuote = new BlockQuote();
                document.Add(blockQuote);
                newBlockQuote = true;
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

                        // Skip any newline/blanks after that.
                        while (i+1 < runs.Count)
                        {
                            if (runs[i + 1].HasText
                                && string.IsNullOrWhiteSpace(runs[i + 1].Text) || runs[i+1].Text == "\n")
                                i++;
                            else break;
                        }
                        break;
                    }

                    newBlockQuote = false;
                }

                if (!processed)
                {
                    var p = blockQuote.LastOrDefault();
                    if (p == null)
                    {
                        p = new Paragraph();
                        blockQuote.Add(p);
                    }
                    
                    // If it's a newline, then just use the empty paragraph.
                    if (run.Text != "\n")
                        renderer.WriteElement(document, p, run, tags);
                    else
                    {
                        p = new Paragraph();
                        blockQuote.Add(p);
                    }
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

            if (document.Last() != blockOwner)
            {
                blockOwner = document.Last();
                if (blockOwner is Paragraph p)
                {

                }
            }
        }

        private static void CreateCodeBlock(IMarkdownRenderer renderer, MarkdownDocument document, 
            MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            Debug.Assert(blockOwner == null);

            string language = null;
            var next = element.NextParagagraph;
            if (next.Properties.StyleName == "CodeFooter")
            {
                language = next.Text;
            }

            var codeBlock = new CodeBlock(language, element.Text);

            // See if we're in a list.
            if (document.Last() is MarkdownList theList && element.Properties.LeftIndent > 0)
            {
                var blocks = theList[^1];
                blocks.Add(codeBlock);
            }
            else document.Add(codeBlock);
        }

        private static void CreateListBlock(IMarkdownRenderer renderer, MarkdownDocument document, MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            var format = element.GetNumberingFormat();

            switch (format)
            {
                case NumberingFormat.Bullet:
                    CreateUnorderedListBlock(renderer, document, blockOwner, element, tags);
                    break;
                case NumberingFormat.None:
                {
                    // Just add to existing list.
                    var list = (MarkdownList) document.Last();
                    var blocks = list[^1];
                    var paragraph = new Paragraph();
                    blocks.Add(paragraph);
                    CreateParagraph(renderer, document, paragraph, element, tags);
                    break;
                }
                // Numbered list
                default:
                    CreateOrderedListBlock(renderer, document, blockOwner, element, tags);
                    break;
            }
        }

        private static void CreateOrderedListBlock(IMarkdownRenderer renderer, MarkdownDocument document, MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            int? level = element.GetListLevel() ?? 0;
            int? index = (element.GetListIndex() ?? 0) + 1;

            // If this is a new list, then create it.
            if (document.Last() is not OrderedList theList)
            {
                theList = new OrderedList(index.Value);
                document.Add(theList);
            }

            blockOwner = new Paragraph();
            theList.Add(blockOwner);

            renderer.WriteContainer(document, blockOwner, element.Runs, tags);
        }

        private static void CreateUnorderedListBlock(IMarkdownRenderer renderer, MarkdownDocument document, MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        {
            int? level = element.GetListLevel() ?? 0;

            // If this is a new list, then create it.
            if (document.Last() is not List theList)
            {
                theList = new List();
                document.Add(theList);
            }

            blockOwner = new Paragraph();
            theList.Add(blockOwner);

            renderer.WriteContainer(document, blockOwner, element.Runs, tags);
        }
    }
}
