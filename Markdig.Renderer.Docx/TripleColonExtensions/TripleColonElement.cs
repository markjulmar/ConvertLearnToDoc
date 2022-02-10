using System.Collections.Generic;
using System.Linq;
using System.Text;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.DocAsCode.MarkdigEngine.Extensions;

namespace Markdig.Renderer.Docx.TripleColonExtensions
{
    public sealed class TripleColonElement
    {
        public ITripleColonExtensionInfo Extension { get; set; }
        public IDictionary<string, string> RenderProperties { get; set; }
        public bool Closed { get; set; }
        public bool EndingTripleColons { get; set; }
        public IDictionary<string, string> Attributes { get; set; }
        public ContainerBlock Container { get; set; }
        public ContainerInline Inlines { get; set; }

        public TripleColonElement(TripleColonBlock block)
        {
            Extension = block.Extension;
            RenderProperties = block.RenderProperties;
            Closed = block.Closed;
            EndingTripleColons = block.EndingTripleColons;
            Attributes = block.Attributes;
            if (block.Count > 0)
            {
                Container = block;
            }
        }

        public TripleColonElement(TripleColonInline inline)
        {
            Extension = inline.Extension;
            RenderProperties = inline.RenderProperties;
            Closed = inline.Closed;
            EndingTripleColons = inline.EndingTripleColons;
            Attributes = inline.Attributes;
            if (inline.Count > 0)
            {
                Inlines = inline;
            }
        }
    }
}