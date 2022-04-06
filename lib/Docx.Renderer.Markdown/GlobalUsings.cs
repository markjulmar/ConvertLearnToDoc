global using DXPlus;
global using Julmar.GenMarkdown;
global using System.Diagnostics;
global using System.Text.RegularExpressions;

global using DXText = DXPlus.Text;
global using DXImage = DXPlus.Image;
global using DXParagraph = DXPlus.Paragraph;
global using DXTable = DXPlus.Table;

global using TableRow = Julmar.GenMarkdown.TableRow;
global using Paragraph = Julmar.GenMarkdown.Paragraph;
global using Text = Julmar.GenMarkdown.Text;

namespace Docx.Renderer.Markdown;

internal static class Globals
{
    public const string UseExtension = "useExtension";
}