namespace Docx.Renderer.Markdown.Renderers;

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
        { "Caption", (_, _, _, _, _) => { /* Do nothing */ } },
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
        {"error", "ERROR"},
        { "important", "IMPORTANT" },
        { "caution","CAUTION"}
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
        if (element.Properties.DefaultFormatting.Subscript)
            tf.Subscript = true;
        if (element.Properties.DefaultFormatting.Superscript)
            tf.Superscript = true;

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
        var p = new Heading(level);
        document.Add(p);
        renderer.WriteContainer(document, p, element.Runs, tags);
    }

    private static void CreateBlockQuote(IMarkdownRenderer renderer, MarkdownDocument document, 
        MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
    {
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
            newBlockQuote = true;

            // See if we're in a list.
            if (document.LastOrDefault() is MarkdownList theList && element.Properties.LeftIndent > 0)
            {
                var blocks = theList[^1];
                blocks.Add(blockQuote);
            }
            else 
                document.Add(blockQuote);
        }

        bool startedText = false;
        var runs = element.Runs.ToList();
        foreach (var run in runs)
        {
            string text = run.Text;
            if (!startedText && (text.Length == 0 || text.All(c => c is ' ' or '\n'))) 
                continue;

            if (newBlockQuote)
            {
                int index = text.IndexOf(':');
                if (index < 0) index = text.IndexOf(' ');

                string key = index < 0 ? text.Trim('\r','\n') : text[..index];
                if (key.Contains(' '))
                {
                    index = text.IndexOf(' ');
                    key = text[..index];
                }

                index = key.Length;

                if (blockHeaders.TryGetValue(key, out string header))
                {
                    newBlockQuote = false;

                    blockQuote.Add($"[!{header}]");
                    blockQuote.Add(new Paragraph());

                    text = text[index..];
                    if (text.FirstOrDefault() == ':')
                        text = text[1..];

                    if (text.Length == 0 || text.All(c => c is ' ' or '\n')) continue;
                }
            }

            newBlockQuote = false;
            startedText = true;

            var p = blockQuote.LastOrDefault();
            if (p == null)
            {
                p = new Paragraph();
                blockQuote.Add(p);
            }
                
            // If it's a newline, then just use the empty paragraph.
            if (text != "\n")
                renderer.WriteElement(document, p, run, tags);
            else
            {
                p = new Paragraph();
                blockQuote.Add(p);
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

        if (element.Text.Contains("{ extended: false }"))
        {

        }

        renderer.WriteContainer(document, blockOwner, element.Runs, tags);
    }

    private static void CreateCodeBlock(IMarkdownRenderer renderer, MarkdownDocument document, 
        MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
    {
        Debug.Assert(blockOwner == null);

        string language = null;
        var next = element.NextParagraph;
        if (next.Properties.StyleName == "CodeFooter")
        {
            language = next.Text;
        }

        var codeBlock = string.IsNullOrEmpty(language) ? new CodeBlock() : new CodeBlock(language);
        codeBlock.Add(element.Text);

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
        var format = element.HasListDetails() 
            ? element.GetNumberingFormat() : NumberingFormat.None;

        switch (format)
        {
            case NumberingFormat.Bullet:
                CreateListBlock<List>(renderer, document, blockOwner, element, tags);
                break;
            case NumberingFormat.Removed:
                CreateParagraph(renderer, document, blockOwner, element, tags);
                break;
            case NumberingFormat.None:
            {
                if (document.LastOrDefault() is MarkdownList list)
                {
                    var blocks = list[^1];
                    var paragraph = new Paragraph();
                    blocks.Add(paragraph);
                    CreateParagraph(renderer, document, paragraph, element, tags);
                }
                else
                {
                    // Default to a bullet list.
                    CreateListBlock<List>(renderer, document, blockOwner, element, tags);
                }
                break;
            }
            // Numbered list
            default:
                CreateListBlock<OrderedList>(renderer, document, blockOwner, element, tags);
                break;
        }
    }

    private static void CreateListBlock<TList>(IMarkdownRenderer renderer, MarkdownDocument document, MarkdownBlock blockOwner, DXParagraph element, RenderBag tags)
        where TList : MarkdownList, new()
    {
        MarkdownList theList = document.Last() as MarkdownList;
        int? index = (element.GetListIndex() ?? 0) + 1;
        int? level = element.GetListLevel() ?? 0;

        // See if we're in a list already.
        if (theList != null)
        {
            while (level > 0)
            {
                if (theList[^1][^1] is MarkdownList checkList)
                {
                    theList = checkList;
                    level--;
                }
                else
                {
                    var list = new TList();
                    if (list is OrderedList ol)
                        ol.StartingNumber = index.Value;
                    theList[^1].Add(list);
                    theList = list;
                    break;
                }
            }
        }

        // If this is a new list, then create it.
        if (theList == null)
        {
            Debug.Assert(level == 0);
            theList = new TList();
            if (theList is OrderedList ol)
                ol.StartingNumber = index.Value;
            document.Add(theList);
        }

        blockOwner = new Paragraph();
        theList.Add(blockOwner);

        renderer.WriteContainer(document, blockOwner, element.Runs, tags);
    }
}