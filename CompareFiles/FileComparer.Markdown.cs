using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace CompareFiles;

public static partial class FileComparer
{
    public static IEnumerable<Difference<MarkdownObject>> Markdown(string fn1, string fn2)
    {
        var pipeline = MarkdownPipeline.Value;

        string text1 = File.ReadAllText(fn1);
        string text2 = File.ReadAllText(fn2);

        var markdownDocument = Markdig.Markdown.Parse(text1, pipeline);
        var compareDocument = Markdig.Markdown.Parse(text2, pipeline);

        int lineNumber = 1;
        using var enumerator = compareDocument.EnumerateBlocks().GetEnumerator();
        foreach (var item in markdownDocument.EnumerateBlocks())
        {
            if (!enumerator.MoveNext())
            {
                yield return new Difference<MarkdownObject>
                {
                    Filename1 = fn1, Filename2 = fn2,
                    LineNumber = lineNumber,
                    Value1 = item, Value2 = null
                };
            }
            else
            {
                var item2 = enumerator.Current;
                if (item.ToString() != item2?.ToString())
                {
                    yield return new Difference<MarkdownObject>
                    {
                        Filename1 = fn1,
                        Filename2 = fn2,
                        LineNumber = lineNumber,
                        Value1 = item,
                        Value2 = item2
                    };
                }
            }

            lineNumber++;
        }
    }

    private static IEnumerable<MarkdownObject> EnumerateBlocks(this ContainerBlock container)
    {
        foreach (var item in container)
        {
            yield return item;
            switch (item)
            {
                case ContainerBlock cb:
                {
                    foreach (var child in EnumerateBlocks(cb))
                        yield return child;
                    break;
                }

                case LeafBlock lb:
                {
                    var containerInline = lb.Inline;
                    if (containerInline != null)
                    {
                        foreach (var child in EnumerateInline(containerInline))
                            yield return child;
                    }
                    break;
                }
            }
        }
    }

    private static IEnumerable<MarkdownObject> EnumerateInline(this Inline inline)
    {
        if (inline is ContainerInline cil)
        {
            foreach (var child in cil)
            {
                yield return child;
                foreach (var sil in EnumerateInline(child))
                    yield return sil;
            }
        }
    }


    private static readonly Lazy<MarkdownPipeline> MarkdownPipeline = new(CreatePipeline);
    private static MarkdownPipeline CreatePipeline()
    {
        var context = new MarkdownContext();
        var pipelineBuilder = new MarkdownPipelineBuilder();
        return pipelineBuilder
            .UsePipeTables()
            .UseRow(context)
            .UseNestedColumn(context)
            .UseTripleColon(context)
            .UseGenericAttributes() // Must be last as it is one parser that is modifying other parsers
            .Build();
    }

}
