namespace Docx.Renderer.Markdown;

public struct TextFormatting
{
    public string StyleName { get; set; }
    public bool Monospace { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool KbdTag { get; set; }
    public bool Superscript { get; set; }
    public bool Subscript { get; set; }

    public static bool IsBoldFont(System.Drawing.FontFamily fontFamily)
    {
        if (fontFamily == null) return false;
        string font = fontFamily.Name.ToLower();
        return font.Contains("bold");
    }

    public static bool IsMonospaceFont(System.Drawing.FontFamily fontFamily)
    {
        if (fontFamily == null) return false;
        string font = fontFamily.Name.ToLower();
        return font.Contains("mono")
               || font.Contains("code")
               || font is "consolas" or "courier new";
    }
}