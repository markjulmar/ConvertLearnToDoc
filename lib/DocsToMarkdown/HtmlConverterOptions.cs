namespace Julmar.DocsToMarkdown;

public class HtmlConverterOptions
{
    public Action<string>? UnhandledTagCallback { get; set; }
    public Action<string, Uri>? DownloadAsset { get; set; }
    public bool UseDocfxExtensions { get; set; }
}