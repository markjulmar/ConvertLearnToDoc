using System.Drawing;

namespace Markdig.Renderer.Docx
{
    internal static class Globals
    {
        public static readonly FontFamily CodeFont = new("Consolas");
        public static readonly Color CodeBoxShade = Color.FromArgb(0xf0, 0xf0, 0xf0);
        public const double CodeFontSize = 10;
    }
}
