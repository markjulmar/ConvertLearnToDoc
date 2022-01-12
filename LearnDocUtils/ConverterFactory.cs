namespace LearnDocUtils
{
    public static class DocxConverterFactory
    {
        public static IDocxToMarkdown WithPandoc => new DocxToMarkdownPandoc();
        public static IDocxToMarkdown WithDxPlus => new DocxToMarkdownDxPlus();
    }

    public static class MarkdownConverterFactory
    {
        public static IMarkdownToDocx WithPandoc => new MarkdownToDocxPandoc();
        public static IMarkdownToDocx WithDxPlus => new MarkdownToDocxDxPlus();
    }
}
