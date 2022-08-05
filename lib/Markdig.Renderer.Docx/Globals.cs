global using DXPlus;
global using System.Text;
global using System.Xml;
global using Markdig.Syntax;
global using Microsoft.DocAsCode.MarkdigEngine.Extensions;
global using Markdig.Syntax.Inlines;

global using Formatting = DXPlus.Formatting;

using System.Drawing;

namespace Markdig.Renderer.Docx;

internal static class Globals
{
    public static readonly FontValue CodeFont = new("Consolas");
    public static readonly Color CodeBoxShade = Color.FromArgb(0xf0, 0xf0, 0xf0);
    public const double CodeFontSize = 10;
    public const string UseExtension = "useExtension";
}