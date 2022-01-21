using System.Linq;
using DXPlus;
using GenMarkdown.DocFx.Extensions;
using Julmar.GenMarkdown;
using DXTable = DXPlus.Table;
using Paragraph = Julmar.GenMarkdown.Paragraph;
using TableRow = Julmar.GenMarkdown.TableRow;

namespace Docx.Renderer.Markdown.Renderers
{
    public sealed class TableRenderer : MarkdownObjectRenderer<DXTable>
    {
        protected override void Render(IMarkdownRenderer renderer, 
            MarkdownDocument document, MarkdownBlock blockOwner, 
            DXTable element, RenderBag tags)
        {
            tags ??= new RenderBag();
            
            var mdTable = new DocfxTable(element.ColumnCount);
            var tcf = element.ConditionalFormatting;
            
            // TODO: pass bold flags down on rows/columns
            var rows = element.Rows.ToList();

            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                var tr = new TableRow();
                mdTable.Add(tr);

                bool boldRow = (rowIndex == 0 && (tcf & TableConditionalFormatting.FirstRow) != 0
                        || rowIndex == rows.Count - 1 && (tcf & TableConditionalFormatting.LastRow) != 0);

                for (var colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                {
                    var col = row.Cells[colIndex];
                    var p = new Paragraph();
                    tr.Add(p);

                    bool boldCol = (colIndex == 0 && (tcf & TableConditionalFormatting.FirstColumn) != 0
                                    || colIndex == row.Cells.Count - 1 &&
                                    (tcf & TableConditionalFormatting.LastColumn) != 0);

                    if (boldRow || boldCol)
                        tags.AddOrReplace(nameof(TextFormatting), new TextFormatting { Bold = true });
                    else 
                        tags.Remove(nameof(TextFormatting));
                    
                    renderer.WriteContainer(document, p, col.Paragraphs, tags);
                }
            }

            document.Add(mdTable);
        }
    }
}
