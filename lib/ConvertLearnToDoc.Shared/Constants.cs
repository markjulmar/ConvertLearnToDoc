namespace ConvertLearnToDoc.Shared;

public static class Constants
{
    public const string WordMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    public const string DefaultModuleName = "learn-module";
    public const string MarkdownExtension = ".md";
    public const string MarkdownMimeType = "text/markdown";
    public const string ZipMimeType = "application/zip";
}

public enum PageType
{
    Article,
    Module,
}