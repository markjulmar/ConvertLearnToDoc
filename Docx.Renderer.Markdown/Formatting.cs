namespace Docx.Renderer.Markdown
{
    public class TextFormatting
    {
        public string StyleName { get; set; }
        public bool Monospace { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
    }

    public class QuoteFormatting
    {
        public string Type { get; set; }
    }

    public class CodeFormatting
    {
        public string Language { get; set; }
    }
}
