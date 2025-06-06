﻿using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using Markdig.Helpers;
using Block = Markdig.Syntax.Block;

namespace Markdig.Renderer.Docx;

public static class MarkdigExtensions
{
    public static IEnumerable<MarkdownObject> EnumerateBlocks(this ContainerBlock container)
    {
        foreach (var item in container)
        {
            yield return item;

            if (item is ContainerBlock cb)
            {
                foreach (var child in EnumerateBlocks(cb))
                    yield return child;
            }
            else if (item is LeafBlock lb)
            {
                var containerInline = lb.Inline;
                if (containerInline != null)
                {
                    foreach (var child in EnumerateInline(containerInline))
                        yield return child;
                }
            }
        }
    }

    public static IEnumerable<MarkdownObject> EnumerateInline(this Inline inline)
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


    public static string Dump(ContainerBlock block, int tabs = 0)
    {
        StringBuilder sb = new StringBuilder();
        DumpContainerBlock(block, tabs, sb);
        return sb.ToString();
    }

    private static void DumpContainerBlock(ContainerBlock cb, int tabs, StringBuilder sb)
    {
        foreach (var item in cb)
        {
            DumpBlock(item, tabs, sb);
        }
    }

    private static void DumpBlock(Block item, int tabs, StringBuilder sb)
    {
        string prefix = new('\t', tabs);

        string details = item switch
        {
            QuoteSectionNoteBlock {VideoLink: { }} qsb => "video: " + qsb.VideoLink,
            QuoteSectionNoteBlock qsb => qsb.NoteTypeString ?? qsb.SectionAttributeString,
            FencedCodeBlock fcb => $"{fcb.Info} - {fcb.Lines.Count} lines",
            TripleColonBlock tcb => $"{tcb.Extension.Name}: {string.Join(", ", tcb.Attributes.Select(a => $"{a.Key}={a.Value}"))}",
            _ => string.Empty
        };

        sb.AppendLine($"{prefix}{item} {details}");

        switch (item)
        {
            case LeafBlock pb:
            {
                var containerInline = pb.Inline;
                if (containerInline != null)
                {
                    foreach (var child in containerInline)
                    {
                        DumpInline(child, tabs + 1, sb);
                    }
                }

                if (pb.Lines.Count > 0)
                {
                    foreach (var str in pb.Lines.Cast<StringLine>().Take(pb.Lines.Count))
                        sb.AppendLine($"{prefix} > {str}");
                }

                break;
            }
            case ContainerBlock cb:
                DumpContainerBlock(cb, tabs + 1, sb);
                break;
        }
    }

    private static void DumpInline(Inline item, int tabs, StringBuilder sb)
    {
        string prefix = new('\t', tabs);

        string typeName = item.GetType().Name;
        string details = GetDebuggerDisplay(item);
        string toString = item.ToString();
        if (details == item.ToString())
            details = "";
        else if (toString == item.GetType().ToString() && item is not ContainerInline)
        {
            toString = details;
            details = "";
        }

        if (item is ContainerInline cil)
        {
            sb.AppendLine($"{prefix}{typeName} {details}");
            DumpContainerInline(cil, tabs + 1, sb);
        }
        else
        {
            sb.AppendLine($"{prefix}{typeName} Value=\"{toString}\" {details}");
        }
    }

    private static void DumpContainerInline(ContainerInline inline, int tabs, StringBuilder sb)
    {
        foreach (var item in inline)
        {
            DumpInline(item, tabs, sb);
        }
    }

    private static string GetDebuggerDisplay(object item)
    {
        var attr = item.GetType().GetCustomAttribute<DebuggerDisplayAttribute>();
        string formatText = attr?.Value;
        if (formatText == null) return null;

        var type = item.GetType();

        foreach (Match match in Regex.Matches(formatText, @"{(.*?)}").Where(m => !m.Groups[1].Value.Contains('{')))
        {
            string pn = match.Groups[1].Value;

            string value;
            var pi = type.GetProperty(pn);
            if (pi != null)
            {
                value = pi.GetValue(item)?.ToString() ?? "(null)";
            }
            else
            {
                value = type.GetField(pn)?.GetValue(item)?.ToString() ?? "(null)";
            }

            formatText = formatText.Replace(match.Value, value);
        }

        return formatText;
    }
}